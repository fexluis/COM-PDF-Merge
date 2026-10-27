using System;
using System.Threading;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Singleton thread-safe responsable de la carga diferida del motor SQLite nativo,
    /// la extensión CR-SQLite, la base de datos académica y el contexto de sincronización.
    /// Evita que el Add-In se demore al arrancar en Excel para usuarios que no usan DB.
    /// </summary>
    internal static class DatabaseBootstrapper
    {
        private static volatile bool _initialized = false;
        private static volatile bool _initFailed = false; // Fase 18
        private static readonly object _initLock = new object();

        /// <summary>
        /// Indica si la base de datos y la sincronización ya han sido inicializadas
        /// exitosamente durante el ciclo de vida actual del Add-in.
        /// </summary>
        internal static bool WasInitialized => _initialized;
        
        internal static bool InitFailed => _initFailed; // Fase 18

        /// <summary>
        /// Restablece el estado para pruebas unitarias.
        /// No debe ser llamado en el flujo de producción.
        /// </summary>
        internal static void ResetStateForTests()
        {
            _initialized = false;
            _initFailed = false; // Fase 18
        }

        /// <summary>
        /// Garantiza que todos los componentes pesados relacionados con la base de datos
        /// (Motor SQLite, CR-SQLite, Migraciones y Sincronización) estén inicializados.
        /// Garantiza la inicialización diferida en dos fases.
        /// FASE A (con progreso): CrSqlite + DB.
        /// FASE B (sin progreso/ProgressManager): SyncContext (puede mostrar diálogos WinForms).
        /// </summary>
        /// <remarks>
        /// PATRÓN DE USO: El llamador en el Ribbon debe invocar este método y esperar
        /// a que retorne ANTES de crear un AcademicDbSession(). AcademicDbSession()
        /// también llama EnsureInitialized() en su constructor como red de seguridad,
        /// pero esa llamada será siempre un no-op si el Ribbon inicializó primero.
        /// </remarks>
        /// <param name="reportStep">
        /// Callback opcional para reportar la etapa actual al llamador.
        /// El llamador puede usarlo para actualizar un ProgressManager antes de
        /// que la Fase B (SyncContext) requiera el hilo UI para sus diálogos.
        /// </param>
        internal static void EnsureInitialized(Action<string> reportStep = null)
        {
            if (_initialized) return;

            lock (_initLock)
            {
                if (_initialized) return;

                // Validación STA: El hilo debe soportar la apertura de ventanas de UI,
                // ya que SyncContext.Initialize() podría mostrar un Formulario WinForms
                // (diálogo de selección de subunidad).
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                {
                    LoggerHelper.Warning("[DatabaseBootstrapper] EnsureInitialized() llamado desde un hilo no-STA. " +
                                         "El diálogo de subunidad no podría mostrarse. " +
                                         "Considerar invocar desde el hilo principal de Excel.");
                    // Permitimos continuar porque podría no requerir el diálogo (ej. usuario ya registrado),
                    // pero queda advertido en caso de InvalidOperationException.
                }

                try
                {
                    LoggerHelper.Info("[DatabaseBootstrapper] Iniciando carga diferida de Base de Datos y Sincronización...");
                    var sw = System.Diagnostics.Stopwatch.StartNew();

                    // --- FASE A: Componentes nativos y DB (sin diálogos WinForms) ---
                    reportStep?.Invoke("Cargando motor de base de datos...");
                    SQLiteEngineBootstrap.EnsureInitialized();

                    reportStep?.Invoke("Extrayendo extensión de sincronización...");
                    CrSqliteLoader.Initialize();

                    reportStep?.Invoke("Restaurando configuración de base de datos...");
                    // Fase 9 FIX: RestoreActiveVaultPath fue refactorizado y movido al pipeline de SyncContext.Initialize()

                    reportStep?.Invoke("Inicializando base de datos académica...");
                    SqliteHelper.GetContext();

                    // --- FASE B: Sync (puede mostrar diálogos WinForms bloqueantes) ---
                    // El llamador debe cerrar el ProgressManager ANTES de que esta
                    // línea se ejecute para que FormRegistroSubunidad sea accesible.
                    reportStep?.Invoke("__CLOSE_PROGRESS__"); // Señal para cerrar el ProgressManager
                    SyncContext.Initialize();

                    sw.Stop();
                    LoggerHelper.Info($"[DatabaseBootstrapper] Carga diferida completada exitosamente en {sw.ElapsedMilliseconds}ms.");

                    _initFailed = false; // Fase 19: Resetear fallo previo
                    _initialized = true;
                }
                catch (Exception ex)
                {
                    _initFailed = true; // Fase 18
                    LoggerHelper.Error("[DatabaseBootstrapper] Error crítico durante la carga diferida de la DB", ex);
                    throw; // Propagamos para que el llamador original (ej. Ribbon) maneje el error visiblemente
                }
                finally
                {
                    // Fase 17: Invalidar siempre, incluso si falló, para salir del fail-open inicial
                    try { ConversorPDF.XLL.RibbonController.InvalidateRibbon(); } catch { }
                }
            }
        }
    }
}
