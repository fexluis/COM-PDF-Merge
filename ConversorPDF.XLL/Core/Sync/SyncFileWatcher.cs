using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Monitorea la carpeta de OneDrive de la subunidad para detectar nuevos archivos
    ///   .crsync entrantes y disparar el PULL automáticamente.
    ///
    /// MECANISMO:
    ///   FileSystemWatcher de .NET detecta eventos Created en la carpeta de sync.
    ///   Cuando OneDrive descarga un archivo .crsync de otra sede, se dispara el evento
    ///   y se procesa el archivo en background (Task.Run) sin bloquear la UI.
    ///
    /// PREVENCIÓN DE DUPLICADOS:
    ///   - Filtra archivos por el patrón "from_*.crsync" (ignore otros archivos).
    ///   - No procesa archivos propios (from_{deviceId}_*).
    ///   - SyncControlRepository verifica sync_archive antes de aplicar.
    ///
    /// RESILIENCIA:
    ///   FileSystemWatcher puede "perder" eventos en carpetas de red bajo alta carga
    ///   (buffer overflow del kernel). El evento Error del watcher detecta esto.
    ///   Cuando ocurre, el watcher se reinicia automáticamente con backoff.
    ///
    /// ARRANQUE:
    ///   Al iniciar el watcher, también procesa archivos .crsync que ya estuvieran
    ///   en la carpeta (para recuperar sync perdido durante downtime del add-in).
    /// </summary>
    internal sealed class SyncFileWatcher : IDisposable
    {
        // ====================================================================
        // CONSTANTES
        // ====================================================================

        /// <summary>
        /// Patrón de archivos a monitorear.
        /// Solo archivos .crsync (ignorar .tmp, .processed, etc.)
        /// </summary>
        private const string FilePattern = "*.crsync";

        /// <summary>
        /// Tiempo máximo de espera en ms antes de leer un archivo recién creado.
        /// OneDrive puede escribir el archivo en múltiples chunks. Esperamos un poco
        /// antes de leerlo para asegurar que la escritura esté completa.
        /// </summary>
        private const int ReadDelayMs = 500;

        // ====================================================================
        // DEPENDENCIAS
        // ====================================================================

        private readonly SubunidadSyncService _syncService;
        private readonly SyncPathResolver _pathResolver;
        private readonly string _subunidad;

        // ====================================================================
        // ESTADO INTERNO
        // ====================================================================

        private FileSystemWatcher _watcher;
        private FileSystemWatcher _globalWatcher;
        private bool _disposed = false;

        /// <summary>
        /// Semáforo para limitar el número de PULLs concurrentes.
        /// Previene que múltiples archivos lleguen al mismo tiempo y saturen la DB.
        /// Máximo 2 PULLs paralelos (suficiente para files en ráfaga).
        /// </summary>
        private readonly SemaphoreSlim _concurrencyLimit = new SemaphoreSlim(2, 2);

        // ====================================================================
        // CONSTRUCTOR
        // ====================================================================

        /// <summary>
        /// Inicializa el watcher para la carpeta OneDrive de la subunidad.
        /// NO inicia el monitoreo aún — llamar Start() para comenzar.
        /// </summary>
        internal SyncFileWatcher(
            SubunidadSyncService syncService,
            SyncPathResolver pathResolver,
            string subunidad)
        {
            _syncService = syncService ?? throw new ArgumentNullException(nameof(syncService));
            _pathResolver = pathResolver ?? throw new ArgumentNullException(nameof(pathResolver));
            _subunidad = subunidad ?? throw new ArgumentNullException(nameof(subunidad));
        }

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Inicia el monitoreo de la carpeta de sync de OneDrive.
        ///
        /// PROCESO:
        ///   1. Resolver y crear la carpeta de sync si no existe.
        ///   2. Procesar archivos .crsync pendientes (downtime recovery).
        ///   3. Iniciar FileSystemWatcher.
        ///
        /// TOLERANCIA A FALLOS:
        ///   Si la carpeta OneDrive no está disponible (sin conexión, no instalado),
        ///   el método loguea y retorna silenciosamente. El sync estará deshabilitado
        ///   hasta que Start() sea llamado de nuevo.
        /// </summary>
        internal void Start()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(SyncFileWatcher));
            }

            try
            {
                // Resolver la carpeta de la subunidad y crearla si no existe
                string syncFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);

                if (!Directory.Exists(syncFolder))
                {
                    Directory.CreateDirectory(syncFolder);
                    LoggerHelper.Info($"[SyncFileWatcher] Carpeta de sync creada: {syncFolder}");
                }

                // Asegurar que existe la carpeta .processed también
                string processedFolder = _pathResolver.GetProcessedFolder(_subunidad);
                if (!Directory.Exists(processedFolder))
                {
                    Directory.CreateDirectory(processedFolder);
                }

                // Iniciar el monitoreo en background para no bloquear el arranque
                // y permitir que el recovery termine antes de activar el watcher.
                Task.Run(async () =>
                {
                    try
                    {
                        string globalFolder = _pathResolver.GetGlobalInboxFolder();
                        if (!Directory.Exists(globalFolder)) Directory.CreateDirectory(globalFolder);
                        string globalProcessed = _pathResolver.GetGlobalProcessedFolder();
                        if (!Directory.Exists(globalProcessed)) Directory.CreateDirectory(globalProcessed);

                        await ProcessPendingFilesOnStartupAsync(syncFolder);
                        InitializeWatcher(syncFolder, ref _watcher);
                        
                        await ProcessPendingFilesOnStartupAsync(globalFolder);
                        InitializeWatcher(globalFolder, ref _globalWatcher);

                        // FASE 9.5: Revisar si hay dispositivos atrasados
                        LoggerHelper.Debug($"[SyncState] Disparando CheckLaggingDevices desde arranque/recovery...");
                        _syncService.CheckLaggingDevices();

                        LoggerHelper.Info($"[SyncFileWatcher] Monitoreo iniciado en: {syncFolder} y {globalFolder}");
                    }
                    catch (Exception ex)
                    {
                        DisposeWatcher();
                        LoggerHelper.Error($"[SyncFileWatcher] Error inicializando: {ex.Message}", ex);
                    }
                });
            }
            catch (DirectoryNotFoundException ex)
            {
                // OneDrive no disponible — sync deshabilitado silenciosamente
                LoggerHelper.Warning(
                    $"[SyncFileWatcher] No se pudo iniciar el monitoreo de sync. " +
                    $"OneDrive no está disponible: {ex.Message}. " +
                    $"El sync PULL estará deshabilitado hasta reiniciar.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SyncFileWatcher] Error al iniciar el monitoreo.", ex);
            }
        }

        /// <summary>
        /// Detiene el monitoreo y libera recursos.
        /// </summary>
        internal void Stop()
        {
            DisposeWatcher();
            LoggerHelper.Info("[SyncFileWatcher] Monitoreo detenido.");
        }

        // ====================================================================
        // INICIALIZACIÓN DEL WATCHER
        // ====================================================================

        /// <summary>
        /// Crea y configura el FileSystemWatcher para la carpeta de sync.
        /// </summary>
        private void InitializeWatcher(string folderPath, ref FileSystemWatcher watcher)
        {
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Created -= OnFileCreated;
                watcher.Error -= OnWatcherError;
                watcher.Dispose();
            }

            watcher = new FileSystemWatcher(folderPath, FilePattern)
            {
                // Monitorear solo archivos en la carpeta raíz, no subcarpetas (.processed)
                IncludeSubdirectories = false,

                // Solo notificar eventos de creación de archivos nuevos
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,

                // Aumentar el buffer interno del watcher para tolerancia a ráfagas
                // El default de 8KB puede perderse si llegan muchos archivos rápidamente
                InternalBufferSize = 65536  // 64 KB
            };

            // Suscribir a eventos
            watcher.Created += OnFileCreated;
            watcher.Error += OnWatcherError;

            // Activar el monitoreo
            watcher.EnableRaisingEvents = true;
        }

        // ====================================================================
        // MANEJADORES DE EVENTOS
        // ====================================================================

        /// <summary>
        /// Se dispara cuando aparece un nuevo archivo .crsync en la carpeta monitorizada.
        ///
        /// CONTEXTO: Este callback corre en un hilo de I/O del kernel (no en Excel STA).
        /// Es seguro iniciar Task.Run aquí.
        ///
        /// DELAY: Esperamos ReadDelayMs antes de procesar para asegurar que OneDrive
        /// terminó de escribir el archivo (puede llegar en chunks).
        /// </summary>
        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            string filePath = e.FullPath;
            string fileName = e.Name;

            LoggerHelper.Debug($"[SyncFileWatcher] Archivo detectado: {fileName}");

            // Validar que es un archivo .crsync (no un .tmp ni otra extensión)
            if (!fileName.EndsWith(".crsync", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            // Lanzar el procesamiento en background con control de concurrencia
            Task.Run(async () =>
            {
                // Esperar a que el archivo esté completamente escrito por OneDrive
                await Task.Delay(ReadDelayMs);

                // Verificar que el archivo sigue existiendo (no fue borrado durante el delay)
                if (!File.Exists(filePath))
                {
                    LoggerHelper.Debug($"[SyncFileWatcher] Archivo ya no existe al procesar: {fileName}");
                    return;
                }

                // Control de concurrencia: esperar si ya hay 2 PULLs en progreso
                await _concurrencyLimit.WaitAsync();
                try
                {
                    await _syncService.ProcessIncomingFileAsync(filePath, CancellationToken.None);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[SyncFileWatcher] Error procesando {fileName}.", ex);
                }
                finally
                {
                    _concurrencyLimit.Release();
                }
            });
        }

        /// <summary>
        /// Se dispara cuando el FileSystemWatcher pierde eventos por buffer overflow.
        /// Esto puede ocurrir en carpetas de red con alta actividad.
        ///
        /// RESPUESTA: Reiniciar el watcher con un delay de backoff para darle tiempo
        /// al sistema a recuperarse. Durante el reinicio, procesamos todos los archivos
        /// pendientes manualmente (recovery).
        /// </summary>
        private void OnWatcherError(object sender, ErrorEventArgs e)
        {
            LoggerHelper.Warning(
                $"[SyncFileWatcher] Error en FileSystemWatcher: {e.GetException()?.Message}. " +
                $"Reiniciando en 5 segundos...");

            // Reiniciar el watcher en un hilo background después de un delay
            Task.Run(async () =>
            {
                await Task.Delay(5000); // backoff de 5 segundos

                // FASE 19: Separar el reinicio de cada watcher en su propio try-catch
                try
                {
                    string syncFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);
                    await ProcessPendingFilesOnStartupAsync(syncFolder);
                    InitializeWatcher(syncFolder, ref _watcher);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[SyncFileWatcher] Error al reiniciar el watcher local.", ex);
                }
                
                try
                {
                    string globalFolder = _pathResolver.GetGlobalInboxFolder();
                    await ProcessPendingFilesOnStartupAsync(globalFolder);
                    InitializeWatcher(globalFolder, ref _globalWatcher);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[SyncFileWatcher] Error al reiniciar el watcher global.", ex);
                }
                
                LoggerHelper.Info("[SyncFileWatcher] Watcher(s) reiniciado(s).");
            });
        }

        // ====================================================================
        // RECOVERY AL ARRANQUE
        // ====================================================================

        /// <summary>
        /// Procesa archivos .crsync que ya existen en la carpeta de sync al arrancar.
        ///
        /// ESCENARIO:
        ///   Si el add-in estuvo apagado (Excel cerrado) mientras otra sede generó
        ///   archivos .crsync, esos archivos estaban esperando en OneDrive.
        ///   Al iniciar el watcher, los procesamos inmediatamente sin esperar nuevos eventos.
        ///
        /// ORDEN: Se procesan en orden alfabético (que es cronológico por el formato del nombre).
        /// </summary>
        private async Task ProcessPendingFilesOnStartupAsync(string syncFolder)
        {
            try
            {
                string[] pendingFiles = Directory.GetFiles(syncFolder, "from_*.crsync");

                if (pendingFiles.Length == 0)
                {
                    LoggerHelper.Debug("[SyncFileWatcher] No hay archivos pendientes al arrancar.");
                    return;
                }

                // Ordenar cronológicamente (el nombre incluye timestamp ordenable)
                Array.Sort(pendingFiles, StringComparer.OrdinalIgnoreCase);

                LoggerHelper.Info(
                    $"[SyncFileWatcher] {pendingFiles.Length} archivo(s) .crsync pendiente(s) al arrancar. " +
                    $"Procesando en recovery...");

                await _concurrencyLimit.WaitAsync();
                try
                {
                    await _syncService.ProcessIncomingFilesBatchAsync(pendingFiles, CancellationToken.None);
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error($"[SyncFileWatcher] Error en batch recovery.", ex);
                }
                finally
                {
                    _concurrencyLimit.Release();
                }

                LoggerHelper.Info("[SyncFileWatcher] Batch recovery completado exitosamente.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncFileWatcher] Error escaneando archivos pendientes: {ex.Message}");
            }
        }

        /// <summary>
        /// Expuesto internamente para permitir que la interfaz gráfica (ej. FormInfo) 
        /// force la verificación asíncrona de archivos pendientes en las carpetas de sincronización.
        /// Esto inicia un ciclo PULL manual de los archivos `.crsync` existentes sin reiniciar los watchers.
        /// </summary>
        internal async Task ForceProcessPendingFilesAsync()
        {
            LoggerHelper.Info("[SyncFileWatcher] Forzando PULL manual de archivos pendientes...");
            
            string syncFolder = _pathResolver.GetSubunidadInboxFolder(_subunidad);
            await ProcessPendingFilesOnStartupAsync(syncFolder);
            
            string globalFolder = _pathResolver.GetGlobalInboxFolder();
            await ProcessPendingFilesOnStartupAsync(globalFolder);
            
            LoggerHelper.Info("[SyncFileWatcher] PULL manual completado.");
        }

        // ====================================================================
        // LIMPIEZA
        // ====================================================================

        private void DisposeWatcher()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Created -= OnFileCreated;
                _watcher.Error -= OnWatcherError;
                _watcher.Dispose();
                _watcher = null;
            }
            if (_globalWatcher != null)
            {
                _globalWatcher.EnableRaisingEvents = false;
                _globalWatcher.Created -= OnFileCreated;
                _globalWatcher.Error -= OnWatcherError;
                _globalWatcher.Dispose();
                _globalWatcher = null;
            }
        }

        public void Dispose()
        {
            if (_disposed) { return; }
            _disposed = true;
            _concurrencyLimit.Dispose();
            DisposeWatcher();
            LoggerHelper.Debug("[SyncFileWatcher] Disposed.");
        }
    }
}
