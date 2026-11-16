using System;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SQLite;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context
{
    /// <summary>
    /// NIVEL 2: Contexto de ciclo de vida para la base de datos SQLite encriptada (Vault Pattern).
    /// </summary>
    public sealed class EncryptedSQLiteContext : IDisposable
    {
        private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, int> _activeContextsPerTempDb = new System.Collections.Concurrent.ConcurrentDictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
        private static readonly string _appRunId = Guid.NewGuid().ToString("N").Substring(0, 8);
        private readonly SecureString _password;
        private readonly string _encryptedFilePath;
        private readonly string _tempDbPath;
        private SQLiteConnection _connection;
        private bool _disposed;
        private bool _connectionIsDisposed;
        private bool _changesPersisted = false; // FIX-0.3: Guardia de integridad
        private bool _isReadOnly = false;
        public string TempDbPath => _tempDbPath;
        public string EncryptedFilePath => _encryptedFilePath;
        public bool IsDisposed => _disposed || _connectionIsDisposed;

        public EncryptedSQLiteContext(string encPath, SecureString password)
        {


            if (string.IsNullOrWhiteSpace(encPath))
                throw new ArgumentNullException(nameof(encPath));
            if (password == null)
                throw new ArgumentNullException(nameof(password));

            // Asegurar que el motor esté listo antes de cualquier operación (Nivel 1)
            SQLiteEngineBootstrap.EnsureInitialized();

            _encryptedFilePath = encPath;
            _password = password.Copy();
            
            // Usamos un ID de ejecución estático para que todos los contextos de la misma DB
            // en la misma sesión compartan el archivo temporal.
            string safeName = Path.GetFileNameWithoutExtension(encPath).Replace(".", "");
            string tempFile = $"esap_acad_tmp_{safeName}_{_appRunId}.db";
            _tempDbPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), tempFile));
            _activeContextsPerTempDb.AddOrUpdate(_tempDbPath, 1, (_, count) => count + 1);
        }

        /// <summary>
        /// Desencripta la base de datos y retorna una conexión SQLite abierta.
        /// </summary>
        public SQLiteConnection Open()
        {
            if (IsDisposed) throw new ObjectDisposedException(nameof(EncryptedSQLiteContext));

            // IDEMPOTENCIA: Reutilizar conexión abierta si existe y es válida (Evita leaks y doble descifrado)
            if (_connection != null 
                && _connection.State == System.Data.ConnectionState.Open 
                && !_connectionIsDisposed)
            {
                LoggerHelper.Info("[EncryptedSQLiteContext] Reutilizando conexión existente.");
                return _connection;
            }

            // Limpiar conexión previa en estado inválido antes de recrear (Paso defensivo)
            if (_connection != null)
            {
                CloseConnectionResiliently();
            }

            LoggerHelper.Info($"[EncryptedSQLiteContext] Iniciando Open() en {_tempDbPath}");

            try
            {
                var lockSw = System.Diagnostics.Stopwatch.StartNew();
                if (!_lock.Wait(TimeSpan.FromSeconds(60)))
                {
                    LoggerHelper.Error($"[EncryptedSQLiteContext] TIMEOUT esperando lock para descifrado. dbPath={_tempDbPath}");
                    throw new TimeoutException("No se pudo obtener acceso al Vault de base de datos (Recurso bloqueado por más de 60s).");
                }
                lockSw.Stop();
                if (lockSw.ElapsedMilliseconds > 1000)
                {
                    LoggerHelper.Warning($"[PERF] [EncryptedSQLiteContext] Alta contención: Esperó {lockSw.ElapsedMilliseconds}ms por el lock de descifrado.");
                }
                
                try
                {
                    if (File.Exists(_encryptedFilePath))
                    {
                        bool needsExtraction = true;
                        if (File.Exists(_tempDbPath))
                        {
                            var tempInfo = new FileInfo(_tempDbPath);
                            if (tempInfo.Length > 0)
                            {
                                needsExtraction = false;
                                LoggerHelper.Debug($"[EncryptedSQLiteContext] Archivo temporal ya existe para esta sesión en {_tempDbPath}. Omitiendo extracción.");
                            }
                        }

                        if (needsExtraction)
                        {
                            // LoggerHelper.Debug($"[EncryptedSQLiteContext] Extrayendo vault (ReadAllBytes)...");
                            byte[] encryptedBlob = File.ReadAllBytes(_encryptedFilePath);
                            
                            // LoggerHelper.Debug($"[EncryptedSQLiteContext] Desencriptando ({encryptedBlob.Length} bytes)...");
                            byte[] plainDatabaseBytes = DatabaseEncryptionEngine.Decrypt(encryptedBlob, _password);
                            
                            File.WriteAllBytes(_tempDbPath, plainDatabaseBytes);
                            
                            // Higiene de memoria (.NET 4.8)
                            Array.Clear(plainDatabaseBytes, 0, plainDatabaseBytes.Length);
                        }
                    }
                    else
                    {
                        if (!File.Exists(_tempDbPath))
                        {
                            SQLiteConnection.CreateFile(_tempDbPath);
                        }
                    }
                }
                finally
                {
                    _lock.Release();
                }

                // Crear conexión con retry para absorber fallos de resolución de ensamblados COM
                _connection = CreateConnectionWithRetry(_tempDbPath);
                _connection.Open();
                
                // Habilitar extensión CR-SQLite en la conexión base para que los triggers funcionen en cualquier contexto
                CrSqliteLoader.EnableOnConnection(_connection);
                
                _connectionIsDisposed = false;
                LoggerHelper.Info($"[EncryptedSQLiteContext] Conexión abierta exitosamente.");
                
                return _connection;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error($"[EncryptedSQLiteContext] Falló la apertura/descifrado del vault", ex);
                CleanupResourcesQuietly();
                throw new Exception($"Error inicializando el contexto encriptado: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Crea una conexión SQLite con reintentos para manejar el bug "path1 null" de System.Data.SQLite
        /// en entornos COM/Excel-DNA.
        /// </summary>
        private static SQLiteConnection CreateConnectionWithRetry(string dbPath, int maxRetries = 2)
        {
            string absolutePath = Path.GetFullPath(dbPath);
            string connString = $"Data Source={absolutePath};Pooling=False;Version=3;";

            for (int i = 0; i < maxRetries; i++)
            {
                try
                {
                    return new SQLiteConnection(connString);
                }
                catch (ArgumentNullException ex) when (ex.ParamName == "path1")
                {
                    if (i == maxRetries - 1) throw;

                    LoggerHelper.Warning($"[EncryptedSQLiteContext] Intento {i + 1} falló por 'path1 null'. Estabilizando AppDomain...");
                }
            }
            return null; // El compilador requiere un retorno, pero el throw superior previene llegar aquí
        }

        /// <summary>
        /// Marca el contexto como persistido exitosamente en el Vault maestro.
        /// Evita la advertencia de prdida de datos durante el Dispose.
        /// </summary>
        public void MarkChangesPersisted() => _changesPersisted = true;

        /// <summary>
        /// Marca el contexto como de solo lectura, suprimiendo la advertencia de pérdida de datos.
        /// </summary>
        public void MarkAsReadOnly() => _isReadOnly = true;



        public void Dispose()
        {
            if (_disposed) return;

            // FIX-0.3, Fase 91 y Fase 163: Avisar si el contexto se cierra sin haber persistido al Vault, solo si hubo mutaciones reales
            bool isTestMode = AppDomain.CurrentDomain.FriendlyName.IndexOf("test", StringComparison.OrdinalIgnoreCase) >= 0;
            if (!_changesPersisted && !_isReadOnly && !isTestMode)
            {
                bool hasChanges = false;
                try
                {
                    if (_connection != null && _connection.State == System.Data.ConnectionState.Open)
                    {
                        using (var cmd = new SQLiteCommand("SELECT total_changes();", _connection))
                        {
                            var result = cmd.ExecuteScalar();
                            if (result != null && int.TryParse(result.ToString(), out int changes))
                            {
                                hasChanges = changes > 0;
                            }
                        }
                    }
                }
                catch { }

                if (hasChanges)
                {
                    LoggerHelper.Warning("[EncryptedSQLiteContext] ADVERTENCIA: Contexto descartado sin persistir al Vault maestro. Posible pérdida de datos.");
                }
            }

            LoggerHelper.Info($"[EncryptedSQLiteContext] Dispose() iniciado.");

            try
            {
                CloseConnectionResiliently();

                // La re-encriptación ahora es DIFERIDA. 
                // Se debe llamar a ReEncryptAndPersist() explícitamente fuera del hilo UI.
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[EncryptedSQLiteContext] Error durante la disposición", ex);
            }
            finally
            {
                int remaining = 0;
                if (_activeContextsPerTempDb.TryGetValue(_tempDbPath, out int currentCount))
                {
                    remaining = currentCount - 1;
                    if (remaining > 0)
                        _activeContextsPerTempDb[_tempDbPath] = remaining;
                    else
                        _activeContextsPerTempDb.TryRemove(_tempDbPath, out _);
                }
                
                // FIX (Error #4): Solo borrar el archivo temporal si somos el último contexto
                // activo en esta sesión, para no borrarle el archivo a un contexto hermano.
                if (remaining <= 0 && File.Exists(_tempDbPath))
                {
                    try { DatabaseEncryptionEngine.SecureWipe(_tempDbPath); } catch { }
                }

                _password.Dispose();
                _disposed = true;
            }
        }

        public void CloseConnectionResiliently(bool isFinalShutdown = false)
        {
            if (_connection == null) return;

            try
            {
                if (!_connectionIsDisposed && _connection.State != System.Data.ConnectionState.Closed)
                {
                    CheckpointWal();
                    _connectionIsDisposed = true; // Prevenir race conditions marcándolo antes de cerrar
                    _connection.Close();
                    _connection.Dispose();
                }
            }
            catch (ObjectDisposedException) { _connectionIsDisposed = true; }
            catch { }

            try { _connection.Dispose(); } catch { }
            _connection = null;
            
            // Liberación de pools de SQLite (solo si es el último contexto o apagado final)
            // REDUCIDO: Solo en final shutdown para evitar contención innecesaria en tests rápidos
            if (isFinalShutdown)
            {
                try { System.Data.SQLite.SQLiteConnection.ClearAllPools(); } catch { }
            }
        }

        /// <summary>
        /// Realiza un backup seguro (online) de la base de datos temporal hacia otro archivo 
        /// sin cerrar ni interrumpir la conexión actual de la UI.
        /// </summary>
        public void BackupTo(string targetPath)
        {
            // Corrección Fase 128: Validar argumento targetPath
            if (string.IsNullOrWhiteSpace(targetPath))
                throw new ArgumentNullException(nameof(targetPath));

            if (_connection == null || _connection.State != System.Data.ConnectionState.Open)
                throw new InvalidOperationException("La conexión debe estar abierta para realizar un backup.");

            // Asegurarnos de que el WAL esté integrado si es posible antes de copiar
            try { CheckpointWal(); } catch { }

            using (var targetConnection = new SQLiteConnection($"Data Source={targetPath};Pooling=False;Version=3;"))
            {
                targetConnection.Open();
                // Copia la base de datos completa ("main" a "main") de una vez (-1 pages)
                _connection.BackupDatabase(targetConnection, "main", "main", -1, null, 0);
            }
            LoggerHelper.Info($"[EncryptedSQLiteContext] Backup completado exitosamente hacia {targetPath}");
        }

        private void CheckpointWal()
        {
            if (_connection == null || _connection.State != System.Data.ConnectionState.Open) return;

            var sw = System.Diagnostics.Stopwatch.StartNew();
            try
            {
                using (var cmd = new SQLiteCommand("PRAGMA wal_checkpoint(TRUNCATE);", _connection))
                {
                    cmd.ExecuteNonQuery();
                }
                LoggerHelper.Info($"[EncryptedSQLiteContext] WAL checkpoint(TRUNCATE) completado en {sw.ElapsedMilliseconds}ms.");
            }
            catch (SQLiteException ex) when (ex.ResultCode == SQLiteErrorCode.Busy)
            {
                LoggerHelper.Warning($"[EncryptedSQLiteContext] WAL checkpoint(TRUNCATE) falló por BUSY. Intentando PASSIVE...");
                try
                {
                    using (var cmd = new SQLiteCommand("PRAGMA wal_checkpoint(PASSIVE);", _connection))
                    {
                        cmd.ExecuteNonQuery();
                    }
                    LoggerHelper.Info($"[EncryptedSQLiteContext] WAL checkpoint(PASSIVE) completado en {sw.ElapsedMilliseconds}ms.");
                }
                catch (SQLiteException ex2)
                {
                    LoggerHelper.Error("[EncryptedSQLiteContext] WAL checkpoint(PASSIVE) también falló.", ex2);
                    // No relanzamos aquí para permitir que Close() intente limpiar, 
                    // pero el riesgo de pérdida de datos en el Vault aumenta.
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[EncryptedSQLiteContext] Error inesperado en WAL checkpoint.", ex);
            }
        }

        /// <summary>
        /// Sobrecarga que recibe los bytes desencriptados en memoria RAM. Evita leer del disco para 
        /// omitir race conditions, guardando directamente al Vault maestro.
        /// </summary>
        public void ReEncryptAndPersist(byte[] plainDatabaseBytes, IProgress<string> progress = null)
        {
            ReEncryptAndPersistStatic(plainDatabaseBytes, _encryptedFilePath, _password, progress);
        }

        /// <summary>Identidad observable del vault; no es una exclusión entre procesos.</summary>
        internal sealed class VaultIdentity
        {
            internal readonly long Length;
            internal readonly DateTime LastWriteUtc;
            internal readonly string Hash;
            internal VaultIdentity(long length, DateTime lastWriteUtc, string hash)
            { Length = length; LastWriteUtc = lastWriteUtc; Hash = hash; }
            internal bool SameAs(VaultIdentity other) => other != null && Length == other.Length
                && LastWriteUtc == other.LastWriteUtc && Hash == other.Hash;
        }

        /// <summary>Descifra una lectura aislada del vault, sin usar el temporal compartido.</summary>
        internal byte[] ReadPublishedBytes() => DatabaseEncryptionEngine.Decrypt(File.ReadAllBytes(_encryptedFilePath), _password);

        /// <summary>Captura tamaño, fecha y SHA-256; rechaza una lectura que cambie durante la captura.</summary>
        internal static VaultIdentity CaptureVaultIdentity(string path)
        {
            if (!File.Exists(path)) return null;
            var before = new FileInfo(path);
            long length = before.Length;
            DateTime modified = before.LastWriteTimeUtc;
            string hash;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var sha = SHA256.Create())
                hash = Convert.ToBase64String(sha.ComputeHash(stream));
            var after = new FileInfo(path);
            if (!after.Exists || length != after.Length || modified != after.LastWriteTimeUtc)
                throw new IOException("VAULT_IDENTIDAD_INESTABLE");
            return new VaultIdentity(length, modified, hash);
        }

        /// <summary>Publica sólo si el vault conserva la identidad del snapshot; confirma bytes cifrados publicados.
        /// La comparación no elimina la ventana TOCTOU entre procesos externos.</summary>
        internal static bool ReEncryptAndPersistIfUnchanged(byte[] bytes, string path, SecureString password,
            VaultIdentity expected, CancellationToken ct)
        {
            bool verified = false;
            ReEncryptAndPersistStatic(bytes, path, password, null, expected, ct, () => verified = true);
            return verified;
        }

        /// <summary>
        /// Versión estática para re-encriptar sin necesidad de instanciar un contexto Dummy.
        /// </summary>
        public static void ReEncryptAndPersistStatic(byte[] plainDatabaseBytes, string vaultPath, SecureString password, IProgress<string> progress = null)
            => ReEncryptAndPersistStatic(plainDatabaseBytes, vaultPath, password, progress, null, CancellationToken.None, null);

        private static void ReEncryptAndPersistStatic(byte[] plainDatabaseBytes, string vaultPath, SecureString password,
            IProgress<string> progress, VaultIdentity expected, CancellationToken ct, Action verified)
        {
            if (plainDatabaseBytes == null || plainDatabaseBytes.Length == 0) 
                throw new ArgumentException("Los bytes de la base de datos no pueden ser nulos o vacíos", nameof(plainDatabaseBytes));
            // Corrección Fase 126: Validar argumento vaultPath
            if (string.IsNullOrWhiteSpace(vaultPath))
                throw new ArgumentNullException(nameof(vaultPath));
            // Corrección Fase 145: Validar argumento password
            if (password == null)
                throw new ArgumentNullException(nameof(password));
            
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            LoggerHelper.Info("[PERF] [EncryptedSQLiteContext] Iniciando ReEncryptAndPersistStatic(RAM)...");
            
            progress?.Report("Cifrando base de datos (Vault)...");
            var swCrypto = System.Diagnostics.Stopwatch.StartNew();
            byte[] newEncryptedBlob = DatabaseEncryptionEngine.Encrypt(plainDatabaseBytes, password);
            swCrypto.Stop();
            LoggerHelper.Info($"[PERF] [EncryptedSQLiteContext] Cifrado de datos: {swCrypto.ElapsedMilliseconds}ms.");
            
            Array.Clear(plainDatabaseBytes, 0, plainDatabaseBytes.Length);

            var lockSw = System.Diagnostics.Stopwatch.StartNew();
            if (!_lock.Wait(TimeSpan.FromSeconds(60)))
            {
                LoggerHelper.Error($"[EncryptedSQLiteContext] TIMEOUT esperando lock para persistencia. activeContexts={_activeContextsPerTempDb.Count}");
                throw new TimeoutException("TIMEOUT: No se pudo obtener acceso exclusivo al Vault para persistencia tras 60 segundos.");
            }
            lockSw.Stop();
            if (lockSw.ElapsedMilliseconds > 1000)
            {
                LoggerHelper.Warning($"[PERF] [EncryptedSQLiteContext] Alta contención: Esperó {lockSw.ElapsedMilliseconds}ms por el lock de persistencia.");
            }

            try
            {
                string newEncPath = vaultPath + ".new";
                progress?.Report("Escribiendo archivo maestro...");
                var swIO = System.Diagnostics.Stopwatch.StartNew();
                File.WriteAllBytes(newEncPath, newEncryptedBlob);
                swIO.Stop();
                LoggerHelper.Info(string.Format(System.Globalization.CultureInfo.InvariantCulture, "[PERF] [EncryptedSQLiteContext] Escritura de archivo maestro: {0}ms.", swIO.ElapsedMilliseconds));

                string bakPath = vaultPath + ".bak";
                ct.ThrowIfCancellationRequested();
                if (verified != null)
                {
                    var current = CaptureVaultIdentity(vaultPath);
                    if (expected == null ? current != null : !expected.SameAs(current))
                        throw new IOException("VAULT_IDENTIDAD_CAMBIO");
                }

                if (File.Exists(vaultPath))
                {
                    try
                    {
                        // File.Replace es atómica a nivel de NTFS: el destino nunca desaparece
                        // y el backup se crea simultáneamente. Previene pérdida de datos en crasheos.
                        File.Replace(newEncPath, vaultPath, File.Exists(bakPath) ? null : bakPath, true);
                    }
                    catch (Exception exReplace)
                    {
                        LoggerHelper.Warning($"[EncryptedSQLiteContext] File.Replace falló (posible FAT32/SMB/Lock parcial). Iniciando fallback manual. Error: {exReplace.Message}");
                        // Un .bak previo no se sacrifica; sin reemplazo atómico no se publica en ese caso.
                        if (File.Exists(bakPath)) throw new IOException("VAULT_BACKUP_PREVIO_PRESERVADO", exReplace);
                        // Si Replace falló después de crear el backup, tampoco se arriesga su pérdida.
                        if (!File.Exists(vaultPath)) throw new IOException("VAULT_REPLACE_ESTADO_INCIERTO", exReplace);
                        File.Move(vaultPath, bakPath);
                        try { File.Move(newEncPath, vaultPath); }
                        catch
                        {
                            File.Move(bakPath, vaultPath); // Restaurar antes de propagar el fallo.
                            throw;
                        }
                    }
                }
                else
                {
                    // Si el vault maestro no existe (primera vez), simplemente movemos el .new
                    File.Move(newEncPath, vaultPath);
                }
                
                if (verified != null)
                {
                    var published = CaptureVaultIdentity(vaultPath);
                    string hash;
                    using (var sha = SHA256.Create()) hash = Convert.ToBase64String(sha.ComputeHash(newEncryptedBlob));
                    if (published == null || published.Length != newEncryptedBlob.Length || published.Hash != hash)
                        throw new IOException("VAULT_POST_PUBLICACION_NO_COINCIDE");
                    verified();
                }
                swTotal.Stop();
                LoggerHelper.Info($"[PERF] [EncryptedSQLiteContext] Re-encriptación RAM completada en {swTotal.ElapsedMilliseconds}ms.");
                progress?.Report("Guardado exitoso.");
            }
            finally
            {
                _lock.Release();
                // Corrección Fase 109: Limpiar archivo intermedio si falló la operación atómica
                try 
                { 
                    string newEncPath = vaultPath + ".new";
                    if (File.Exists(newEncPath)) File.Delete(newEncPath); 
                } 
                catch { }
            }
        }

        /// <summary>
        /// Lee los bytes de un archivo temporal con esquema de reintento exponencial (Backoff).
        /// Útil en entornos background donde antivirus o indexadores pueden bloquear transitoriamente el archivo.
        /// </summary>
        public static async Task<byte[]> LeerBytesConRetryStatic(string tempPath, IProgress<string> progress = null, CancellationToken ct = default)
        {
            if (!File.Exists(tempPath)) return null;

            int intentos = 0;
            const int maxIntentos = 5;
            while (intentos < maxIntentos)
            {
                ct.ThrowIfCancellationRequested();

                try
                {
                    using (var fs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        byte[] buffer = new byte[fs.Length];
                        fs.Read(buffer, 0, buffer.Length);
                        return buffer;
                    }
                }
                catch (IOException) when (intentos < maxIntentos - 1)
                {
                    intentos++;
                    int delay = 100 * (1 << intentos); // 200, 400, 800, 1600...
                    progress?.Report($"Acceso bloqueado por el sistema, reintentando en {delay}ms...");
                    
                    try { await Task.Delay(delay, ct).ConfigureAwait(false); } 
                    catch (TaskCanceledException) { throw; }
                }
            }
            throw new IOException($"No se pudo leer el archivo temporal en {tempPath} tras agotar {maxIntentos} reintentos.");
        }

        /// <summary>
        /// Versión síncrona de LeerBytesConRetry para contextos donde el async/await no es viable (legacy initialization).
        /// </summary>
        public static byte[] LeerBytesConRetrySync(string tempPath, IProgress<string> progress = null)
        {
            if (!File.Exists(tempPath)) return null;

            int intentos = 0;
            const int maxIntentos = 5;
            while (intentos < maxIntentos)
            {
                try
                {
                    using (var fs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        byte[] buffer = new byte[fs.Length];
                        fs.Read(buffer, 0, buffer.Length);
                        return buffer;
                    }
                }
                catch (IOException) when (intentos < maxIntentos - 1)
                {
                    intentos++;
                    int delay = 100 * (1 << intentos); 
                    progress?.Report($"[SYNC] Acceso bloqueado, reintentando en {delay}ms...");
                    System.Threading.Thread.Sleep(delay);
                }
            }
            throw new IOException($"No se pudo leer el archivo temporal (SYNC) en {tempPath} tras {maxIntentos} intentos.");
        }



        private void CleanupResourcesQuietly()
        {
            try { if (File.Exists(_tempDbPath)) DatabaseEncryptionEngine.SecureWipe(_tempDbPath); } catch { }
        }

        private static DateTime _lastCleanupTime = DateTime.MinValue;
        private static readonly TimeSpan CleanupThrottle = TimeSpan.FromHours(1);
        private const string CleanupMutexName = "Local\\ConversorPDF_Cleanup_Mutex";

        internal static void CleanupOrphanedTempFiles(TimeSpan? minAge = null, bool force = false)
        {
            // THROTTLE LOCAL (AppDomain): Evitar llamadas innecesarias en el mismo contexto de carga
            if (!force && DateTime.Now - _lastCleanupTime < CleanupThrottle)
            {
                return;
            }
            _lastCleanupTime = DateTime.Now;

            // THROTTLE GLOBAL (Proceso): Sincronizar entre múltiples AppDomains de Excel-DNA
            using (var mutex = new Mutex(false, CleanupMutexName))
            {
                bool acquired = false;
                try
                {
                    // Si otro AppDomain ya está limpiando o limpió hace poco, omitimos.
                    // Timeout de 0 para no bloquear hilos UI/Ribbon si hay contención.
                    // Corrección Fase 133: Capturar AbandonedMutexException si otro proceso crasheó
                    try
                    {
                        acquired = mutex.WaitOne(0);
                    }
                    catch (AbandonedMutexException)
                    {
                        acquired = true;
                    }
                    
                    if (!acquired)
                    {
                        return;
                    }

                    var threshold = minAge ?? TimeSpan.FromHours(1);
                    string tempPath = Path.GetTempPath();
                    var orphaned = Directory.GetFiles(tempPath, "esap_acad_tmp_*.db");
                    
                    if (orphaned.Length > 0)
                    {
                        LoggerHelper.Info($"[EncryptedSQLiteContext] Iniciando limpieza de {orphaned.Length} archivos temporales (Sincronizado via Mutex)...");
                    }

                    int deletedCount = 0;
                    int ignoredCount = 0;

                    foreach (var file in orphaned)
                    {
                        try
                        {
                            var fi = new FileInfo(file);
                            if (DateTime.Now - fi.LastWriteTime < threshold)
                            {
                                ignoredCount++;
                                continue;
                            }

                            File.Delete(file);
                            deletedCount++;
                        }
                        catch (IOException) 
                        { 
                            ignoredCount++; 
                        }
                        catch (UnauthorizedAccessException) 
                        { 
                            ignoredCount++; 
                        }
                        catch { /* Omitir otros fallos transitorios */ }
                    }

                    if (orphaned.Length > 0)
                    {
                        LoggerHelper.Info($"[EncryptedSQLiteContext] Limpieza finalizada: {deletedCount} eliminados, {ignoredCount} ignorados/en uso.");
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[EncryptedSQLiteContext] Error durante limpieza sincronizada de temporales", ex);
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }
    }
}

