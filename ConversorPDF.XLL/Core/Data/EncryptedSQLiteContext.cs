using System;
using System.IO;
using System.Security;
using System.Threading;
using System.Threading.Tasks;
using System.Data.SQLite;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Data
{
    /// <summary>
    /// NIVEL 2: Contexto de ciclo de vida para la base de datos SQLite encriptada (Vault Pattern).
    /// </summary>
    public sealed class EncryptedSQLiteContext : IDisposable
    {
        private static readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
        private readonly SecureString _password;
        private readonly string _encryptedFilePath;
        private readonly string _tempDbPath;
        private SQLiteConnection _connection;
        private bool _disposed;
        private bool _connectionIsDisposed;
        private bool _changesPersisted = false; // FIX-0.3: Guardia de integridad
        public string TempDbPath => _tempDbPath;

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
            
            string tempFile = "esap_acad_tmp_" + Guid.NewGuid().ToString("N") + ".db";
            _tempDbPath = Path.GetFullPath(Path.Combine(Path.GetTempPath(), tempFile));
        }

        /// <summary>
        /// Desencripta la base de datos y retorna una conexión SQLite abierta.
        /// </summary>
        public SQLiteConnection Open()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(EncryptedSQLiteContext));

            LoggerHelper.Info($"[EncryptedSQLiteContext] Iniciando Open() en {_tempDbPath}");

            try
            {
                if (!_lock.Wait(TimeSpan.FromSeconds(30)))
                {
                    throw new TimeoutException("No se pudo obtener acceso al Vault de base de datos (Recurso bloqueado).");
                }
                
                try
                {
                    if (File.Exists(_encryptedFilePath))
                    {
                        byte[] encryptedBlob = File.ReadAllBytes(_encryptedFilePath);
                        byte[] plainDatabaseBytes = DatabaseEncryptionEngine.Decrypt(encryptedBlob, _password);
                        
                        File.WriteAllBytes(_tempDbPath, plainDatabaseBytes);
                        
                        // Higiene de memoria (.NET 4.8)
                        Array.Clear(plainDatabaseBytes, 0, plainDatabaseBytes.Length);
                    }
                    else
                    {
                        SQLiteConnection.CreateFile(_tempDbPath);
                    }
                }
                finally
                {
                    _lock.Release();
                }

                // Crear conexión con retry para absorber fallos de resolución de ensamblados COM
                _connection = CreateConnectionWithRetry(_tempDbPath);
                _connection.Open();
                
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
            const int retryDelayMs = 50; // Delay corto para no bloquear Excel unnecessarily
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

        public void Dispose()
        {
            if (_disposed) return;

            // FIX-0.3: Avisar si el contexto se cierra sin haber persistido al Vault
            if (!_changesPersisted)
            {
                LoggerHelper.Warning("[EncryptedSQLiteContext] ADVERTENCIA: Contexto descartado sin persistir al Vault maestro. Posible prdida de datos.");
            }

            LoggerHelper.Info($"[EncryptedSQLiteContext] Dispose() iniciado.");

            try
            {
                CloseConnectionResiliently();

                // La re-encriptación ahora es DIFERIDA. 
                // Se debe llamar a ReEncryptAndPersist() explícitamente fuera del hilo UI.
                
                if (File.Exists(_tempDbPath))
                {
                    DatabaseEncryptionEngine.SecureWipe(_tempDbPath);
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[EncryptedSQLiteContext] Error durante la disposición", ex);
            }
            finally
            {
                _password.Dispose();
                _disposed = true;
            }
        }

        public void CloseConnectionResiliently()
        {
            if (_connection == null) return;

            try
            {
                if (!_connectionIsDisposed && _connection.State != System.Data.ConnectionState.Closed)
                {
                    _connection.Close();
                }
            }
            catch (ObjectDisposedException) { _connectionIsDisposed = true; }
            catch { }

            try { _connection.Dispose(); } catch { }
            _connection = null;
            
            // Liberación de pools de SQLite (necesario para liberar el archivo temporal)
            try { System.Data.SQLite.SQLiteConnection.ClearAllPools(); } catch { }
        }

        /// <summary>
        /// Sobrecarga que recibe los bytes desencriptados en memoria RAM. Evita leer del disco para 
        /// omitir race conditions, guardando directamente al Vault maestro.
        /// </summary>
        public void ReEncryptAndPersist(byte[] plainDatabaseBytes, IProgress<string> progress = null)
        {
            ReEncryptAndPersistStatic(plainDatabaseBytes, _encryptedFilePath, _password, progress);
        }

        /// <summary>
        /// Versión estática para re-encriptar sin necesidad de instanciar un contexto Dummy.
        /// </summary>
        public static void ReEncryptAndPersistStatic(byte[] plainDatabaseBytes, string vaultPath, SecureString password, IProgress<string> progress = null)
        {
            if (plainDatabaseBytes == null || plainDatabaseBytes.Length == 0) 
                throw new ArgumentException("Los bytes de la base de datos no pueden ser nulos o vacíos", nameof(plainDatabaseBytes));
            
            var swTotal = System.Diagnostics.Stopwatch.StartNew();
            LoggerHelper.Info("[PERF] [EncryptedSQLiteContext] Iniciando ReEncryptAndPersistStatic(RAM)...");
            
            progress?.Report("Cifrando base de datos (Vault)...");
            var swCrypto = System.Diagnostics.Stopwatch.StartNew();
            byte[] newEncryptedBlob = DatabaseEncryptionEngine.Encrypt(plainDatabaseBytes, password);
            swCrypto.Stop();
            LoggerHelper.Info($"[PERF] [EncryptedSQLiteContext] Cifrado de datos: {swCrypto.ElapsedMilliseconds}ms.");
            
            Array.Clear(plainDatabaseBytes, 0, plainDatabaseBytes.Length);

            if (!_lock.Wait(TimeSpan.FromSeconds(30)))
            {
                throw new TimeoutException("TIMEOUT: No se pudo obtener acceso exclusivo al Vault para persistencia tras 30 segundos.");
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

                if (File.Exists(vaultPath))
                {
                    try
                    {
                        // File.Replace es atómica a nivel de NTFS: el destino nunca desaparece
                        // y el backup se crea simultáneamente. Previene pérdida de datos en crasheos.
                        File.Replace(newEncPath, vaultPath, bakPath, true);
                    }
                    catch (Exception exReplace)
                    {
                        LoggerHelper.Warning($"[EncryptedSQLiteContext] File.Replace falló (posible FAT32/SMB/Lock parcial). Iniciando fallback manual. Error: {exReplace.Message}");
                        // Fallback manual para sistemas de archivos que no soporten File.Replace (FAT32, SMB, etc.)
                        if (File.Exists(bakPath)) File.Delete(bakPath);
                        File.Move(vaultPath, bakPath);
                        File.Move(newEncPath, vaultPath);
                    }
                }
                else
                {
                    // Si el vault maestro no existe (primera vez), simplemente movemos el .new
                    File.Move(newEncPath, vaultPath);
                }
                
                swTotal.Stop();
                LoggerHelper.Info($"[PERF] [EncryptedSQLiteContext] Re-encriptación RAM completada en {swTotal.ElapsedMilliseconds}ms.");
                progress?.Report("Guardado exitoso.");
            }
            finally
            {
                _lock.Release();
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
                    using (var fs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
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
                    
                    try { await Task.Delay(delay, ct); } 
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
                    using (var fs = new FileStream(tempPath, FileMode.Open, FileAccess.Read, FileShare.Read))
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

        internal static void CleanupOrphanedTempFiles()
        {
            try
            {
                string tempPath = Path.GetTempPath();
                var orphaned = Directory.GetFiles(tempPath, "esap_acad_tmp_*.db");
                foreach (var file in orphaned)
                {
                    try
                    {
                        // Intentar borrar. Si está en uso por otra instancia de Excel, fallará y es normal.
                        File.Delete(file);
                        LoggerHelper.Info($"[EncryptedSQLiteContext] Archivo temporal huérfano eliminado: {Path.GetFileName(file)}");
                    }
                    catch { /* En uso */ }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[EncryptedSQLiteContext] Error durante limpieza de temporales", ex);
            }
        }
    }
}

