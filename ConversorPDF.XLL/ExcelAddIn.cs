using System;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;
using ConversorPDF.XLL.Core.Sync;
using ExcelDna.Integration;

namespace ConversorPDF.XLL
{
    /// <summary>
    /// Punto de entrada del Add-in. Se ejecuta al cargar/descargar desde Excel.
    ///
    /// FASE 1 — PER-APPDOMAIN (se ejecuta en cada AppDomain al cargar el complemento):
    ///   1a. SQLiteEngineBootstrap  — motor SQLite nativo (SQLite.Interop.dll)
    ///   1b. Registro de manejadores de eventos no controlados.
    ///   1c. Registro de limpieza al salir del proceso (ProcessExit).
    ///
    /// FASE 2 — PER-PROCESS (se ejecuta solo una vez por instancia de excel.exe):
    ///   2.  Marca de variable de entorno "CONVERSORPDF_BOOTSTRAPPED".
    ///
    /// NOTA: La inicialización de la Base de Datos (CrSqliteLoader, SqliteHelper, SyncContext)
    /// ya NO se realiza aquí. Ahora se realiza bajo demanda mediante
    /// DatabaseBootstrapper.EnsureInitialized().
    /// </summary>
    public class ExcelAddIn : IExcelAddIn
    {
        private static int _appDomainInitState = 0; // 0 = no ejecutado, 1 = en progreso, 2 = completado
        private static readonly System.Threading.Mutex _autoOpenGlobalMutex = new System.Threading.Mutex(
            false,
            $"Local\\ConversorPDF_AutoOpen_{Environment.UserName}_{Environment.MachineName}");

        public void AutoOpen()
        {
            // ===================================================================
            // FASE 1: PER-APPDOMAIN (Se ejecuta siempre, en todos los AppDomains)
            // ===================================================================
            
            // Guardia intra-dominio: evita re-entrada dentro del mismo AppDomain
            if (System.Threading.Interlocked.CompareExchange(ref _appDomainInitState, 1, 0) == 0)
            {
                try
                {
                    // 1a. Motor SQLite nativo (idempotente por dominio)
                    SQLiteEngineBootstrap.EnsureInitialized();

                    // 1b. Manejadores de eventos locales al dominio
                    ExcelIntegration.RegisterUnhandledExceptionHandler(ex =>
                    {
                        var exception = ex as Exception;
                        LoggerHelper.Error("Error no controlado capturado por ExcelIntegration", exception);

                        // Limpieza de emergencia
                        SQLiteEngineBootstrap.EnsureInitialized();

                        return $"[ConversorPDF] Error no controlado: {exception?.Message ?? ex?.ToString()}";
                    });

                    // 1c. Registro para limpieza al salir del proceso
                    AppDomain.CurrentDomain.ProcessExit += (s, e) =>
                    {
                        EncryptedSQLiteContext.CleanupOrphanedTempFiles(TimeSpan.FromHours(1), force: false);
                    };
                }
                finally
                {
                    System.Threading.Interlocked.Exchange(ref _appDomainInitState, 2);
                }
            }

            // ===================================================================
            // FASE 2: PER-PROCESS GLOBAL (Se ejecuta solo UNA vez por excel.exe)
            // ===================================================================

            // Fast Path: Otro AppDomain ya completó la fase global
            if (Environment.GetEnvironmentVariable("CONVERSORPDF_BOOTSTRAPPED", EnvironmentVariableTarget.Process) == "1")
            {
                return;
            }

            bool mutexAcquired = false;
            try
            {
                // Guardia INTER-APPDOMAIN de exclusión mutua
                try
                {
                    mutexAcquired = _autoOpenGlobalMutex.WaitOne(0);
                }
                catch (System.Threading.AbandonedMutexException)
                {
                    // El AppDomain que tenía el mutex murió inesperadamente.
                    // AbandonedMutexException en WaitOne significa que LO ADQUIRIMOS de todas formas.
                    mutexAcquired = true;
                }

                if (!mutexAcquired)
                {
                    // Otro AppDomain está ejecutando la fase global ahora. 
                    return;
                }

                // Doble comprobación de idempotencia dentro de la sección crítica
                if (Environment.GetEnvironmentVariable("CONVERSORPDF_BOOTSTRAPPED", EnvironmentVariableTarget.Process) == "1")
                {
                    return;
                }

                var swTotal = System.Diagnostics.Stopwatch.StartNew();
                LoggerHelper.Info("==================================================================");
                LoggerHelper.Info("[ExcelAddIn] INICIANDO ARRANQUE DEL COMPLEMENTO v" + typeof(ExcelAddIn).Assembly.GetName().Version);
                LoggerHelper.Info("==================================================================");

                // Marcar proceso completado (indica que el Add-in arrancó, NO que la DB se cargó)
                Environment.SetEnvironmentVariable("CONVERSORPDF_BOOTSTRAPPED", "1", EnvironmentVariableTarget.Process);

                // Lanzar la inicialización pesada de la BD en un hilo STA de background.
                // Se requiere STA (en lugar de Task.Run/ThreadPool) porque SyncContext.Initialize()
                // puede mostrar formularios WinForms (FormRegistroSubunidad) que necesitan un
                // hilo STA para crear ventanas. Al completar, EnsureInitialized() invoca
                // InvalidateRibbon() en su bloque finally, forzando a Excel a re-evaluar
                // GetButtonVisible con los permisos reales del usuario.
                var staInitThread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        DatabaseBootstrapper.EnsureInitialized();
                    }
                    catch (Exception exInit)
                    {
                        LoggerHelper.Error("[ExcelAddIn] Error en inicialización de BD en background", exInit);
                        // EnsureInitialized ya llama InvalidateRibbon() en su finally incluso
                        // si falla, de modo que el Ribbon saldrá del estado fail-closed inicial.
                    }
                });
                staInitThread.SetApartmentState(System.Threading.ApartmentState.STA);
                staInitThread.IsBackground = true;
                staInitThread.Name = "ConversorPDF_DBInit_STA";
                staInitThread.Start();

                swTotal.Stop();
                LoggerHelper.Info($"[ExcelAddIn] Arranque ligero completado en {swTotal.ElapsedMilliseconds}ms. La DB se cargará en background (hilo STA).");

            }
            finally
            {
                if (mutexAcquired)
                {
                    _autoOpenGlobalMutex.ReleaseMutex();
                }
            }
        }

        public void AutoClose()
        {
            // Detener el sistema de sync y forzar persistencia del Vault local
            // Solo ejecutar si la DB llegó a inicializarse en la sesión actual
            if (DatabaseBootstrapper.WasInitialized)
            {
                if (SyncContext.Scheduler != null)
                {
                    try
                    {
                        var sharedCtx = SqliteHelper.GetContext();
                        if (sharedCtx != null)
                        {
                            // Se ejecuta sincrónicamente porque AutoClose no es async en Excel-DNA
                            SyncContext.Scheduler.PersistAndShutdownAsync(sharedCtx).GetAwaiter().GetResult();
                        }
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Error("[ExcelAddIn] Error en PersistAndShutdownAsync durante AutoClose", ex);
                    }
                }
            }

            // Limpiar archivos temporales de SQLite huérfanos
            EncryptedSQLiteContext.CleanupOrphanedTempFiles(TimeSpan.FromHours(1), force: true);
        }
    }
}