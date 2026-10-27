using System;
using System.Threading;
using System.Threading.Tasks;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using System.IO;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Orquesta el disparo de PUSH de forma no-bloqueante para el hilo principal de Excel.
    ///
    /// PATRÓN: Write-Behind + Debounce.
    ///   1. Los repositorios llaman MarkDirty() después de cada INSERT/UPDATE/DELETE.
    ///   2. MarkDirty() es O(1): solo setea un flag y reinicia un timer.
    ///   3. Cuando el timer expira (2 segundos de inactividad), se dispara PushChangesAsync
    ///      en un hilo del ThreadPool (Task.Run), NO en el hilo de Excel.
    ///   4. La UI de Excel NUNCA espera al sync.
    ///
    /// DEBOUNCE:
    ///   Si el usuario hace 10 cambios en 1 segundo, el timer se reinicia 10 veces.
    ///   El PUSH solo ocurre una vez, 2 segundos después del ÚLTIMO cambio.
    ///   Esto agrupa los cambios de operaciones batch (ej: importar 100 terceros).
    ///
    /// PREVENCIÓN DE REENTRANCIA:
    ///   Si llega un MarkDirty() mientras ya hay un PUSH en ejecución:
    ///   - El flag _isDirty se setea a true.
    ///   - Al terminar el PUSH actual, detecta _isDirty y programa otro PUSH.
    ///   - Nunca hay dos PUSH simultáneos.
    ///
    /// EVENTOS:
    ///   SyncStarted y SyncCompleted se disparan en el hilo background.
    ///   Si la UI necesita actualizarse, debe usar Invoke() para volver al STA thread.
    /// </summary>
    internal sealed class SyncScheduler : IDisposable
    {
        // ====================================================================
        // CONFIGURACIÓN
        // ====================================================================

        /// <summary>
        /// Tiempo de espera en milisegundos después del último cambio antes de hacer PUSH.
        /// 2 segundos es el valor recomendado: suficientemente corto para ser "casi inmediato",
        /// suficientemente largo para agrupar cambios batch.
        /// </summary>
        internal int DebounceMs { get; set; } = 5000;

        /// <summary>
        /// Permite deshabilitar temporalmente el scheduler sin destruirlo.
        /// Útil durante importaciones masivas donde queremos acumular todos los cambios
        /// y hacer un único PUSH al final:
        ///
        ///   scheduler.Enabled = false;
        ///   // importar 10000 filas...
        ///   scheduler.Enabled = true;
        ///   scheduler.MarkDirty(); // único PUSH
        /// </summary>
        internal bool Enabled { get; set; } = true;

        // ====================================================================
        // ESTADO INTERNO (thread-safe)
        // ====================================================================

        // Estado para timer adaptativo
        private int _consecutiveDirtyCalls = 0;
        private DateTime _lastDirtyCall = DateTime.MinValue;

        // Estado para coalescencia de PUSH pequeños (Thread-safe lock-free)
        private long _lastPushBytes = 0;
        private long _lastPushTimeTicks = DateTime.MinValue.Ticks;

        /// <summary>
        /// true si hay cambios pendientes de sincronizar.
        /// volatile garantiza que el cambio sea visible entre hilos sin necesitar lock.
        /// </summary>
        private volatile bool _isDirty = false;

        /// <summary>
        /// true si hay un PUSH en progreso en el ThreadPool.
        /// Previene que se acumulen múltiples tareas de PUSH simultáneas.
        /// </summary>
        private volatile bool _isSyncing = false;

        /// <summary>
        /// Fase 10 FIX: Exponer estado de sync para evitar bloqueos (database is locked) en guardados manuales.
        /// </summary>
        internal bool IsSyncing => _isSyncing;

        /// <summary>
        /// Timer de disparo único (no periódico).
        /// Se configura con period = Timeout.Infinite para que no repita.
        /// Cada llamada a MarkDirty() reinicia el timer desde cero.
        /// </summary>
        private Timer _debounceTimer;

        /// <summary>
        /// Servicio de sync inyectado. Contiene la lógica real de PUSH.
        /// </summary>
        private readonly SubunidadSyncService _syncService;

        /// <summary>
        /// Lock para operaciones de doble escritura (_isDirty + timer reinicio).
        /// El lock es pequeño y rápido (solo protege cambios de flags booleanos).
        /// </summary>
        private readonly object _stateLock = new object();

        /// <summary>
        /// Flag de disposed para IDisposable.
        /// </summary>
        private volatile bool _disposed = false;

        /// <summary>
        /// Indica si hay cambios locales pendientes de ser enviados.
        /// </summary>
        internal bool IsDirty
        {
            get
            {
                lock (_stateLock) { return _isDirty; }
            }
        }

        // ====================================================================
        // EVENTOS
        // ====================================================================

        /// <summary>
        /// Se dispara cuando una operación de PUSH comienza (en hilo background).
        /// Útil para mostrar indicador "Sincronizando..." en la barra de estado del Ribbon.
        /// </summary>
        internal event EventHandler SyncStarted;

        /// <summary>
        /// Se dispara cuando una operación de PUSH termina (en hilo background).
        /// El EventArgs contiene el resultado (éxito, bytes enviados, error).
        /// </summary>
        internal event EventHandler<SyncCompletedEventArgs> SyncCompleted;

        // ====================================================================
        // CONSTRUCTOR
        // ====================================================================

        /// <summary>
        /// Inicializa el scheduler con el servicio de sync para una subunidad.
        ///
        /// NOTA: El timer se crea en estado "detenido" (dueTime = Timeout.Infinite).
        /// Solo se activa cuando MarkDirty() lo reinicia.
        /// </summary>
        /// <param name="syncService">Servicio que ejecuta el PUSH real.</param>
        internal SyncScheduler(SubunidadSyncService syncService)
        {
            _syncService = syncService ?? throw new ArgumentNullException(nameof(syncService));

            // Crear timer en estado "dormido": no se dispara hasta MarkDirty()
            _debounceTimer = new Timer(
                callback: OnDebounceElapsed,
                state: null,
                dueTime: Timeout.Infinite,   // No empieza hasta MarkDirty()
                period: Timeout.Infinite);   // No se repite (disparo único)

            LoggerHelper.Info($"[SyncScheduler] Inicializado. DebounceMs: {DebounceMs}ms");
        }

        // ====================================================================
        // API PÚBLICA
        // ====================================================================

        /// <summary>
        /// Marca que hay cambios pendientes y programa el próximo PUSH.
        ///
        /// COMPORTAMIENTO:
        ///   1. Si Enabled = false → retorna inmediatamente sin hacer nada.
        ///   2. Setea _isDirty = true (indica que hay cambios).
        ///   3. Reinicia el timer de debounce a {DebounceMs}ms.
        ///      Si ya había un timer corriendo, se reinicia desde cero (debounce real).
        ///   4. Retorna en &lt; 1 microsegundo. No hace I/O ni operaciones lentas.
        ///
        /// DÓNDE LLAMAR:
        ///   En cada repositorio, después de INSERT/UPDATE/DELETE sobre tablas sincronizables:
        ///
        ///   connection.Execute("INSERT INTO terceros ...");
        ///   SyncContext.Scheduler?.MarkDirty();  // ← única línea agregada
        ///
        /// THREAD SAFETY: Puede llamarse desde cualquier hilo.
        /// </summary>
        internal void MarkDirty()
        {
            if (!Enabled)
            {
                LoggerHelper.Debug("[SyncScheduler] MarkDirty ignorado (scheduler deshabilitado).");
                return;
            }

            int currentConsecutive;
            int adaptiveMs;
            lock (_stateLock)
            {
                if (_disposed) return;
                
                _isDirty = true;

                // Timer adaptativo: aumentar el debounce si el usuario está en una ráfaga de cambios
                double elapsedMs = (DateTime.Now - _lastDirtyCall).TotalMilliseconds;
                if (elapsedMs < 2000)
                {
                    _consecutiveDirtyCalls++;
                }
                else
                {
                    _consecutiveDirtyCalls = 0;
                }
                _lastDirtyCall = DateTime.Now;

                // Escala suave: cada 2 llamadas consecutivas suma un múltiplo
                // Con DebounceMs=5000: 0-1 calls→5s, 2-3→10s, 4-5→15s (techo)
                int multiplier = 1 + (_consecutiveDirtyCalls / 2);
                adaptiveMs = Math.Min(DebounceMs * multiplier, 15000);
                
                // Coalescencia: si el último PUSH fue reciente y muy pequeño (< 1KB), 
                // dar tiempo a acumular más cambios para evitar saturar el flujo de I/O en la red.
                // IMPORTANTE: Solo aplicar si estamos en una ráfaga (isBurst)
                bool isBurst = _consecutiveDirtyCalls > 0;
                long lastBytes = Interlocked.Read(ref _lastPushBytes);
                long lastTicks = Interlocked.Read(ref _lastPushTimeTicks);
                DateTime lastPushTime = new DateTime(lastTicks, DateTimeKind.Utc);

                if (isBurst && lastBytes > 0 && lastBytes < 1024 && (DateTime.UtcNow - lastPushTime).TotalSeconds < 30)
                {
                    adaptiveMs = 30000;
                    LoggerHelper.Debug("[SyncScheduler] MarkDirty: coalescencia de ráfaga pequeña, debounce extendido a 30s.");
                }
                
                currentConsecutive = _consecutiveDirtyCalls;

                try
                {
                    _debounceTimer?.Change(adaptiveMs, Timeout.Infinite);
                }
                catch (ObjectDisposedException)
                {
                    LoggerHelper.Debug("[SyncScheduler] Timer disposed en MarkDirty.");
                }
            }

            // Fase 595: una sola línea por ráfaga (1ª llamada y cada 100), sin log por llamada
            if (currentConsecutive == 1 || currentConsecutive % 100 == 0)
            {
                LoggerHelper.Debug(
                    $"[SyncScheduler] MarkDirty(): pendientes={currentConsecutive}, próximo PUSH en {adaptiveMs}ms.");
            }
        }

        /// <summary>
        /// Fuerza el envío inmediato de los cambios pendientes si los hay.
        /// Cancela el timer actual para no generar un PUSH duplicado.
        /// 
        /// THREAD SAFETY: Seguro llamar desde el hilo STA (OnFormClosing).
        /// NO hace await — dispara el PUSH en el ThreadPool y retorna inmediatamente.
        /// 
        /// USO: Llamar desde OnFormClosing de formularios que realizan escrituras.
        /// Complementa (NO reemplaza) el debounce automático de MarkDirty().
        /// </summary>
        internal void Flush()
        {
            bool shouldSync = false;

            lock (_stateLock)
            {
                if (_disposed) return;
                
                // Detener timer para no generar PUSH duplicado tras el flush
                try { _debounceTimer?.Change(Timeout.Infinite, Timeout.Infinite); }
                catch (ObjectDisposedException) { return; }

                if (_isDirty && !_isSyncing)
                {
                    _isSyncing = true;
                    _isDirty = false;
                    shouldSync = true;
                }
            }

            if (!shouldSync)
            {
                LoggerHelper.Debug("[SyncScheduler] Flush(): sin cambios pendientes o sync en progreso.");
                return;
            }

            LoggerHelper.Info("[SyncScheduler] Flush(): Iniciando PUSH inmediato en background...");

            Task.Run(async () =>
            {
                try
                {
                    SyncStarted?.Invoke(this, EventArgs.Empty);
                    PushResult result = await _syncService.PushChangesAsync(CancellationToken.None);
                    SyncCompleted?.Invoke(this, new SyncCompletedEventArgs(
                        success: true, bytesSent: result.BytesSent,
                        changesCount: result.ChangesCount, error: null));
                    if (result.HadChanges)
                    {
                        LoggerHelper.Info(
                            $"[SyncScheduler] Flush() completado. " +
                            $"Cambios: {result.ChangesCount}, Bytes: {result.BytesSent:N0}, Archivo: {result.FileName}");

                        // FIX: Auto-persistir al Vault después de un Flush exitoso (Fase 189)
                        var sharedCtx = SqliteHelper.TryGetSharedContextWithoutInit();
                        if (sharedCtx != null)
                        {
                            await ForcePersistVaultAsync(sharedCtx).ConfigureAwait(false);
                        }
                        else
                        {
                            LoggerHelper.Debug("[SyncScheduler] Flush(): contexto no disponible; se persistirá desde UI.");
                        }
                    }
                    else
                    {
                        LoggerHelper.Debug("[SyncScheduler] Flush(): Sin cambios pendientes.");
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[SyncScheduler] Error en Flush().", ex);
                    SyncCompleted?.Invoke(this, new SyncCompletedEventArgs(
                        success: false, bytesSent: 0, changesCount: 0, error: ex));
                }
                finally
                {
                    lock (_stateLock)
                    {
                        _isSyncing = false;
                        if (!_disposed && _isDirty && _debounceTimer != null)
                        {
                            try { _debounceTimer.Change(DebounceMs, Timeout.Infinite); }
                            catch (ObjectDisposedException) { }
                        }
                    }
                }
            });
        }

        /// <summary>
        /// Espera de forma bloqueante hasta que el PUSH actual termine, o hasta que se alcance el timeout.
        /// Seguro llamar desde Dispose para no perder la última sincronización.
        /// </summary>
        /// <param name="timeout">Tiempo máximo de espera.</param>
        internal void WaitForPendingSync(TimeSpan timeout)
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.Elapsed < timeout)
            {
                lock (_stateLock)
                {
                    if (!_isSyncing) return;
                }
                Thread.Sleep(100);
            }
            LoggerHelper.Warning("[SyncScheduler] WaitForPendingSync: timeout esperando PUSH.");
        }

        // ====================================================================
        // CICLO DE PUSH INTERNO DEL TIMER (ThreadPool)
        // ====================================================================

        /// <summary>
        /// Callback invocado por el timer cuando han pasado {DebounceMs}ms sin nuevos cambios.
        ///
        /// IMPORTANTE: Este método corre en un hilo del ThreadPool (NO en Excel STA thread).
        /// Es seguro hacer operaciones de I/O aquí sin bloquear la UI.
        ///
        /// ANTI-REENTRANCIA:
        ///   - Si _isSyncing = true (otro PUSH en progreso): no hace nada, el PUSH actual
        ///     se encargará de relanzar si _isDirty sigue siendo true al terminar.
        ///   - Si _isDirty = false: no hay cambios que sync, salir.
        ///
        /// NOTA: El callback es un delegado void (requerido por Timer).
        ///   Internamente usa Task.Run para lanzar la operación async.
        ///   NO usamos async void porque puede causar problemas de manejo de excepciones.
        /// </summary>
        private void OnDebounceElapsed(object state)
        {
            bool shouldSync = false;

            lock (_stateLock)
            {
                if (_disposed) return;

                // Solo proceder si hay cambios Y no hay otro PUSH en progreso
                if (_isDirty && !_isSyncing)
                {
                    _isSyncing = true;   // Bloquear nuevos PUSH concurrentes
                    _isDirty = false;    // Limpiar flag (se re-setea si llegan más cambios)
                    shouldSync = true;
                }
            }

            if (!shouldSync)
            {
                LoggerHelper.Debug(
                    $"[SyncScheduler] OnDebounceElapsed: saltando (isDirty={_isDirty}, isSyncing={_isSyncing})");
                return;
            }

            LoggerHelper.Info("[SyncScheduler] Iniciando PUSH en background...");

            // Lanzar el PUSH en el ThreadPool.
            // NO esperamos el Task (fire and forget con manejo de error en el propio Task).
            // La razón: este callback es sync (Timer no soporta async callbacks en .NET 4.8).
            Task.Run(async () =>
            {
                try
                {
                    // Notificar inicio (para el Ribbon puede mostrar ícono de sync)
                    SyncStarted?.Invoke(this, EventArgs.Empty);

                    // Ejecutar el PUSH real
                    PushResult result = await _syncService.PushChangesAsync(CancellationToken.None);

                    // Notificar resultado
                    SyncCompleted?.Invoke(this, new SyncCompletedEventArgs(
                        success: true,
                        bytesSent: result.BytesSent,
                        changesCount: result.ChangesCount,
                        error: null));

                    if (result.HadChanges)
                    {
                        // Estadísticas lock-free para la coalescencia
                        Interlocked.Exchange(ref _lastPushBytes, result.BytesSent);
                        Interlocked.Exchange(ref _lastPushTimeTicks, DateTime.UtcNow.Ticks);

                        LoggerHelper.Info(
                            $"[SyncScheduler] PUSH exitoso. " +
                            $"Cambios: {result.ChangesCount}, Bytes: {result.BytesSent:N0}, " +
                            $"Archivo: {result.FileName}");

                        // FIX: Auto-persistir al Vault después de un PUSH exitoso (Fase 189)
                        var sharedCtx = SqliteHelper.TryGetSharedContextWithoutInit();
                        if (sharedCtx != null)
                        {
                            await ForcePersistVaultAsync(sharedCtx).ConfigureAwait(false);
                        }
                        else
                        {
                            LoggerHelper.Debug("[SyncScheduler] PUSH: contexto no disponible; se persistirá desde UI.");
                        }
                    }
                    else
                    {
                        LoggerHelper.Debug("[SyncScheduler] PUSH: Sin cambios pendientes.");
                    }
                }
                catch (Exception ex)
                {
                    // Capturar excepciones del PUSH y loguear sin crashear el add-in
                    LoggerHelper.Error("[SyncScheduler] Error durante PUSH en background.", ex);
                    SyncCompleted?.Invoke(this, new SyncCompletedEventArgs(
                        success: false, bytesSent: 0, changesCount: 0, error: ex));
                }
                finally
                {
                    lock (_stateLock)
                    {
                        _isSyncing = false;
                        if (!_disposed && _isDirty && _debounceTimer != null)
                        {
                            try
                            {
                                _debounceTimer.Change(DebounceMs, Timeout.Infinite);
                            }
                            catch (ObjectDisposedException)
                            {
                                LoggerHelper.Debug("[SyncScheduler] Timer disposed. Ignorando reprogramación.");
                            }
                        }
                    }
                }
            });
        }

        // ====================================================================
        // IDISPOSABLE
        // ====================================================================

        /// <summary>
        /// Detiene el scheduler y espera a que el PUSH en progreso termine (máximo 10 segundos).
        ///
        /// Se debe llamar desde ExcelAddIn.AutoClose() para asegurar que el último
        /// PUSH pendiente se complete antes de cerrar el add-in.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) { return; }
            _disposed = true;

            // Detener el timer inmediatamente
            _debounceTimer?.Dispose();
            _debounceTimer = null;

            // Esperar al PUSH en progreso (máximo 10 segundos para no bloquear el cierre de Excel)
            int maxWaitMs = 10000;
            int waitedMs = 0;
            const int pollIntervalMs = 100;

            // Fase 14 FIX: leer _isSyncing bajo lock para evitar stale read concurrente con finally blocks
            while (waitedMs < maxWaitMs)
            {
                lock (_stateLock)
                {
                    if (!_isSyncing) break;
                }
                Thread.Sleep(pollIntervalMs);
                waitedMs += pollIntervalMs;
            }

            bool timedOut;
            lock (_stateLock) { timedOut = _isSyncing; }
            if (timedOut)
            {
                LoggerHelper.Warning("[SyncScheduler] Timeout esperando al PUSH en progreso durante Dispose(). El sync puede haberse interrumpido.");
            }
            else
            {
                LoggerHelper.Info("[SyncScheduler] Scheduler detenido limpiamente.");
            }
        }

        // ====================================================================
        // PERSISTENCIA SEGURA
        // ====================================================================

        /// <summary>
        /// Persiste el TempDB local al Vault maestro usando un backup online sin cerrar la conexión compartida de la UI.
        /// </summary>
        public async Task ForcePersistVaultAsync(EncryptedSQLiteContext sharedCtx)
        {
            if (sharedCtx == null) throw new ArgumentNullException(nameof(sharedCtx));
            try
            {
                if (sharedCtx.IsDisposed)
                {
                    LoggerHelper.Warning("[SyncScheduler] El contexto compartido está cerrado/dispuesto. Se omitirá el backup.");
                    return;
                }
                
                // FIX: Asegurar explícitamente que la conexión a SQLite esté abierta antes 
                // de ejecutar el `BackupTo`. Esto evita la excepción de conexión cerrada si el
                // ciclo de vida de la UI la cerró temporalmente, garantizando que el vault 
                // local persista sus últimos cambios antes de cerrar la aplicación.
                sharedCtx.Open();

                string backupPath = sharedCtx.TempDbPath + ".backup";
                sharedCtx.BackupTo(backupPath);

                // ConfigureAwait(false): LeerBytesConRetryStatic es I/O, no necesita
                // volver al hilo STA. Evita deadlock si se llama con GetAwaiter().GetResult().
                byte[] bytes = await EncryptedSQLiteContext.LeerBytesConRetryStatic(backupPath).ConfigureAwait(false);
                if (bytes != null)
                {
                    SqliteHelper.UpdateVaultDirectly(bytes, sharedCtx.EncryptedFilePath);
                    sharedCtx.MarkChangesPersisted();
                    LoggerHelper.Info("[SyncScheduler] Vault local persistido exitosamente en background.");
                }

                // Usar Thread.Sleep aquí es seguro porque este método es async Task y siempre
                // se invoca desde Task.Run (ThreadPool). El await anterior (línea ~540) ya
                // devolvió el control; la continuación corre en el pool, NO en el hilo STA.
                // Se evita Task.Delay para no introducir un await adicional que complique
                // la ruta de llamada sincrónica de PersistAndShutdownAsync.
                for (int i = 0; i < 3; i++)
                {
                    try { File.Delete(backupPath); break; }
                    catch (IOException) when (i < 2) { Thread.Sleep(500); }
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SyncScheduler] Error persistiendo vault en background", ex);
            }
        }

        /// <summary>
        /// Fuerza la persistencia del Vault y el apagado del Scheduler (Cierre de aplicación).
        /// </summary>
        public async Task PersistAndShutdownAsync(EncryptedSQLiteContext sharedCtx)
        {
            LoggerHelper.Info("[SyncScheduler] PersistAndShutdownAsync iniciado...");
            
            // NOTA: Este método siempre se llama desde el hilo STA mediante
            // GetAwaiter().GetResult(). Usamos Thread.Sleep (no Task.Delay) para
            // evitar deadlock con el SynchronizationContext de WinForms.
            int maxWaitMs = 10000;
            int waitedMs = 0;
            // Fase 17 FIX: usar variable local 'timedOut' para preservar el resultado del loop.
            // Leer _isSyncing post-loop fuera de lock puede dar stale read si el PUSH
            // terminó exactamente al salir del while.
            bool timedOut = true;
            while (waitedMs < maxWaitMs)
            {
                lock (_stateLock)
                {
                    if (!_isSyncing) { timedOut = false; break; }
                }
                Thread.Sleep(100);
                waitedMs += 100;
            }
            
            if (timedOut)
            {
                LoggerHelper.Warning("[SyncScheduler] Timeout esperando PUSH. El background persistirá el vault.");
            }
            else if (sharedCtx != null)
            {
                await ForcePersistVaultAsync(sharedCtx).ConfigureAwait(false);
            }
            
            Dispose();
        }
    }
}
