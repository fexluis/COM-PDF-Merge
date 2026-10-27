using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Linq;
using System.Diagnostics;
using ConversorPDF.XLL.Core.Helpers;
using ExcelDna.Integration;
using System.Data.SQLite;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context
{
    /// <summary>
    /// NIVEL 1: Inicialización Única por PROCESO (Idempotencia Global).
    /// Maneja el estado estático del motor SQLite y la resolución de dependencias nativas.
    /// Utiliza un Mutex Local y verificación de módulos cargados para evitar redundancia
    /// incluso cuando Excel-DNA carga el assembly en múltiples contextos/AppDomains.
    /// </summary>
    public static class SQLiteEngineBootstrap
    {
        private static readonly string MutexName = $"ConversorPDF_SQLiteInit_{Environment.UserName}_{Environment.MachineName}";
        
        // Cache local por AppDomain (CPU-friendly)
        private static bool _localInitialized = false;

        // Cache del resultado de verificación de módulos (para evitar iterar Process.Modules repetidamente)
        private static bool? _interopLoadedCache = null;

        /// <summary>
        /// Asegura que el motor SQLite esté configurado y listo para usarse.
        /// Este método es idempotente y thread-safe a nivel de proceso.
        /// </summary>
        /// <remarks>
        /// CRÍTICO — Comportamiento de Excel-DNA / Shadow Copy:
        /// Excel-DNA carga los assemblies mediante Shadow Copy en rutas temporales tipo
        /// %TEMP%\{guid}\assembly\dl3\{hash}\System.Data.SQLite.dll.
        /// Esto causa que cada contexto de carga tenga su propia copia de los campos estáticos
        /// (como <c>static bool _initialized</c>), haciendo que los guards de campo estático
        /// fallen como singletons de proceso.
        ///
        /// SOLUCIÓN: Usamos <see cref="Process.GetCurrentProcess()"/>.Modules que refleja el
        /// estado real del proceso (excel.exe) independientemente de cuántos AppDomains o
        /// shadow-copies existan. Esto provee idempotencia verdadera a nivel de proceso.
        ///
        /// Un <see cref="Mutex"/> con nombre local de sesión sincroniza la inicialización
        /// entre todos los contextos concurrentes.
        ///
        /// NO intentes reemplazar la lógica de módulos con campos estáticos. Ver evidencia:
        /// docs/log_20260417.txt — 18 ejecuciones redundantes antes de este fix.
        /// Ver: .agents/rules/sqlite-excel-dna-gotchas.md
        /// </remarks>
        internal static void EnsureInitialized()
        {
            // 1. Bypass ultra-rápido: ya se inicializó en este AppDomain
            if (_localInitialized) return;

            // 2. Bypass global: ya se inicializó en otro AppDomain/contexto del mismo proceso
            if (IsInteropAlreadyLoaded())
            {
                _localInitialized = true;
                return;
            }

            // 3. Coordinación inter-AppDomain (Cross-context lock)
            // Usamos namespace Local (sin prefijo Global\) para evitar problemas de permisos de administrador.
            using (var mutex = new Mutex(false, MutexName))
            {
                bool acquired = false;
                try
                {
                    // FASE 3++: Timeout generoso (15s) para colas de AppDomains en máquinas lentas/corporativas.
                    acquired = mutex.WaitOne(TimeSpan.FromSeconds(15), false);
                    
                    if (!acquired)
                    {
                        LoggerHelper.Warning("[SQLiteEngineBootstrap] Timeout esperando Mutex de inicialización (15s).");
                        
                        // VERIFICACIÓN POST-TIMEOUT: ¿Alguien más lo logró?
                        if (IsInteropAlreadyLoaded())
                        {
                            LoggerHelper.Info("[SQLiteEngineBootstrap] Motor detectado en memoria tras el timeout. Continuando de forma segura bajo el supuesto de que otro contexto terminó la carga.");
                            _localInitialized = true;
                            return;
                        }

                        LoggerHelper.Error("[SQLiteEngineBootstrap] Timeout del mutex Y motor NO detectado en memoria. Estado inconsistente/Deadlock.");
                        throw new InvalidOperationException(
                            "El motor SQLite no pudo inicializarse dentro del tiempo esperado (15s) y no se detecta cargado en el proceso. " +
                            "Esto suele indicar un bloqueo en la carga de SQLite.Interop.dll o un conflicto de recursos. " +
                            "Por favor, reinicie Excel.");
                    }

                    // Doble verificación después de adquirir el lock con éxito
                    if (IsInteropAlreadyLoaded())
                    {
                        _localInitialized = true;
                        return;
                    }

                    RealInitialization();
                    _localInitialized = true;
                }
                catch (Exception ex)
                {
                    LoggerHelper.Error("[SQLiteEngineBootstrap] Fallo catastrófico en coordinación Mutex", ex);
                }
                finally
                {
                    if (acquired) mutex.ReleaseMutex();
                }
            }
        }

        /// <summary>
        /// Verifica si la DLL nativa SQLite.Interop.dll ya está mapeada en la memoria del proceso.
        /// </summary>
        /// <returns><c>true</c> si la DLL ya está cargada; <c>false</c> en caso contrario.</returns>
        /// <remarks>
        /// Usa <see cref="Process.GetCurrentProcess()"/>.Modules porque es la única fuente de
        /// verdad que sobrevive a múltiples AppDomains y shadow-copies. Los campos estáticos
        /// son por-copia-del-assembly y NO son confiables en el contexto de Excel-DNA.
        /// El resultado se cachea en <see cref="_interopLoadedCache"/> para evitar enumerar
        /// módulos en cada llamada (O(n) con n ≈ 100-200 DLLs en excel.exe).
        /// </remarks>
        private static bool IsInteropAlreadyLoaded()
        {
            if (_interopLoadedCache.HasValue) return _interopLoadedCache.Value;

            try
            {
                // Enumerar módulos cargados en el proceso actual (excel.exe)
                // Esta es la ÚNICA fuente de verdad global en el proceso.
                var process = Process.GetCurrentProcess();
                _interopLoadedCache = process.Modules
                    .Cast<ProcessModule>()
                    .Any(m => m.ModuleName.IndexOf("SQLite.Interop", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            catch (Exception ex)
            {
                // En algunos entornos restringidos, acceder a Modules puede fallar. Fallback a false.
                LoggerHelper.Warning($"[SQLiteEngineBootstrap] Falló verificación de módulos: {ex.Message}");
                return false;
            }

            return _interopLoadedCache.Value;
        }

        private static void RealInitialization()
        {
            LoggerHelper.Info("[SQLiteEngineBootstrap] RealInitialization() ejecutándose UNA vez desde este AppDomain...");
            
            try
            {
                // 1. Asegurar archivo .config dummy para evitar bug "path1 null" en Shadow Copy
                EnsureSQLiteConfigurationFile();

                // 2. Resolver y pre-caragar SQLite.Interop.dll (v1.2.50.1 Hotfix)
                LoggerHelper.Info($"[Boot] Aplicación v{App.APP_VERSION} | Opt: Shared Context | Engine: v1.0.119.0");
                NativeLibraryPreloader.PreloadSQLite();

                // *** 3. NUEVO: Warm-up del subsistema SQLiteLog (Fase 3) ***
                // RAZÓN CRÍTICA: El constructor de INSTANCIA de `new SQLiteConnection()`
                // invoca `SQLiteLog.Initialize(className)` en cada instancia creada.
                // La primera llamada en el proceso ejecuta `PrivateInitialize()`, que hace
                // el P/Invoke bloqueante `sqlite3_config_log`. Si este primer llamado ocurre
                // dentro de los candados de Fase 2 (_autoOpenGlobalMutex + _lock), el pump de
                // mensajes del hilo STA de Excel no puede avanzar -> Deadlock perfecto y silencioso.
                //
                // SOLUCIÓN: Crear y abrir una conexión real contra :memory: AQUÍ, en la Fase 1,
                // fuera de cualquier candado. Esto fuerza que SQLiteLog.Initialize() y su
                // P/Invoke se ejecuten libres de cualquier contención.
                // La segunda llamada (en EncryptedSQLiteContext.Open) encontrará SQLiteLog
                // ya inicializado y pasará instantáneamente.
                try
                {
                    using (var warmupConn = new System.Data.SQLite.SQLiteConnection("Data Source=:memory:;Version=3;"))
                    {
                        warmupConn.Open();
                        // La conexión se cierra y destruye aquí, sin efectos secundarios.
                    }
                    LoggerHelper.Info("[SQLiteEngineBootstrap] Warm-up de SQLiteLog completado via conexión :memory:.");
                }
                catch (Exception exWarmup)
                {
                    // No es fatal: registramos la anomalía y continuamos.
                    // Indica un problema con SQLite.Interop.dll que debería investigarse.
                    LoggerHelper.Warning($"[SQLiteEngineBootstrap] Warm-up de SQLiteLog falló: {exWarmup.Message}");
                }

                // FASE 3: Guard de Carga de Assembly (Prevenir path1 null)
                AppDomain.CurrentDomain.AssemblyResolve += OnAssemblyResolve;

                LoggerHelper.Info("[SQLiteEngineBootstrap] Motor SQLite estabilizado globalmente para el proceso.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SQLiteEngineBootstrap] Error crítico en RealInitialization", ex);
            }
        }


        /// <summary>
        /// Crea un archivo .config vacío en el shadow-copy directory de System.Data.SQLite.dll.
        /// Requerido para prevenir ArgumentNullException("path1") en el constructor de SQLiteConnection.
        /// </summary>
        private static void EnsureSQLiteConfigurationFile()
        {
            try
            {
                var assembly = typeof(SQLiteConnection).Assembly;
                string assemblyPath = assembly.Location;
                if (string.IsNullOrEmpty(assemblyPath)) return;

                string configPath = assemblyPath + ".config";
                if (!File.Exists(configPath))
                {
                    File.WriteAllText(configPath, "<?xml version=\"1.0\" encoding=\"utf-8\"?><configuration></configuration>");
                    LoggerHelper.Info($"[SQLiteEngineBootstrap] Config dummy creado en shadow-copy: {configPath}");
                }
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SQLiteEngineBootstrap] No se pudo crear config dummy: {ex.Message}");
            }
        }

        private static Assembly OnAssemblyResolve(object sender, ResolveEventArgs args)
        {
            if (args.Name.Contains("System.Data.SQLite"))
            {
                LoggerHelper.Info($"[Boot] Resolviendo Assembly: {args.Name}...");
                var assembly = AppDomain.CurrentDomain.GetAssemblies()
                    .FirstOrDefault(a => a.FullName == args.Name);
                if (assembly != null) return assembly;

                // Si no está cargado, intentar cargar desde la ruta del XLL
                try
                {
                    string xllDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath);
                    string path = Path.Combine(xllDir, "System.Data.SQLite.dll");
                    if (File.Exists(path))
                    {
                        LoggerHelper.Info($"[Boot] Cargando System.Data.SQLite desde: {path}");
                        return Assembly.LoadFrom(path);
                    }
                }
                catch { /* Ignorar fallos en resolución custom */ }
            }
            return null;
        }
    }
}
