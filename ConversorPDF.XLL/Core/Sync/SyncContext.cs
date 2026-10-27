using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Punto central de acceso al sistema de sincronización.
    ///   Actúa como Composition Root del subsistema de sync dentro del add-in.
    ///
    /// PATRÓN SINGLETON:
    ///   Esta clase mantiene las instancias del scheduler y el watcher como
    ///   propiedades estáticas. El add-in siempre tiene UNA subunidad activa
    ///   por sesión de Excel, por lo que un singleton es apropiado.
    ///
    /// CICLO DE VIDA:
    ///   1. Initialize() → llamado desde ExcelAddIn.AutoOpen()
    ///   2. Scheduler?.MarkDirty() → llamado desde repositorios tras CRUD
    ///   3. Shutdown() → llamado desde ExcelAddIn.AutoClose()
    ///
    /// ACCESO DESDE REPOSITORIOS:
    ///   Los repositorios acceden al scheduler via:
    ///     SyncContext.Scheduler?.MarkDirty();
    ///   El operador ?. asegura que si el sync no está habilitado (usuario no registrado,
    ///   OneDrive no disponible), la llamada no falla.
    ///
    /// THREAD SAFETY:
    ///   Initialize() y Shutdown() solo se llaman desde el hilo principal de Excel.
    ///   MarkDirty() es thread-safe internamente (SyncScheduler usa lock).
    /// </summary>
    internal static class SyncContext
    {
        // ====================================================================
        // SINGLETON STATE
        // ====================================================================

        /// <summary>
        /// Nombre del archivo sentinela que indica que el usuario optó
        /// por no participar en la sincronización. Creado cuando el
        /// usuario cancela el diálogo de registro.
        /// </summary>
        private const string SyncOptOutFileName = ".sync_optout";

        /// <summary>
        /// Scheduler de PUSH. null si el sync está deshabilitado.
        /// Los repositorios llaman Scheduler?.MarkDirty() después de cada CRUD.
        /// </summary>
        public static SyncScheduler Scheduler { get; private set; }

        /// <summary>
        /// Watcher de PULL. null si el sync está deshabilitado.
        /// </summary>
        private static SyncFileWatcher _watcher;

        /// <summary>
        /// Subunidad activa en esta sesión de Excel.
        /// null si el sync está deshabilitado.
        /// </summary>
        public static string ActiveSubunidad { get; private set; }

        /// <summary>
        /// true si el sistema de sync fue inicializado exitosamente.
        /// </summary>
        public static bool IsEnabled { get; private set; }
        
        /// <summary>
        /// true si el sistema de sync se está inicializando actualmente en segundo plano.
        /// </summary>
        public static bool IsInitializing => _syncInitState == 1;

        private static int _syncInitState = 0; // 0=pendiente, 1=en progreso, 2=completado

        // ====================================================================
        // COMPOSITION ROOT
        // ====================================================================

        /// <summary>
        /// Inicializa el sistema de sync completo.
        ///
        /// PRERREQUISITO: SqliteHelper.GetContext() debe haber sido llamado antes
        /// (y haber ejecutado las migraciones, incluida V24).
        ///
        /// PROCESO:
        ///   1. Verificar que CR-SQLite esté disponible (ya inicializado).
        ///   2. Resolver la subunidad del usuario actual.
        ///   3. Si el usuario no está registrado → mostrar diálogo bloqueante.
        ///   4. Si el usuario cancela → sync deshabilitado silenciosamente.
        ///   5. Si OK → crear Scheduler + FileWatcher.
        ///
        /// IMPORTANTE: Este método se llama desde el hilo principal de Excel (STA).
        /// El diálogo WinForms es compatible con el STA thread.
        /// </summary>
        internal static void Initialize()
        {
            // Guardia atómica: si ya está en progreso o completado, salir.
            if (System.Threading.Interlocked.CompareExchange(ref _syncInitState, 1, 0) != 0)
                return;

            bool taskLaunched = false;
            try
            {
                // ================================================================
                // PASO 0: Verificar sentinela de exclusión voluntaria
                // ================================================================
                if (IsSyncOptedOut())
                {
                    LoggerHelper.Info("[SyncContext] Sync deshabilitado por exclusión voluntaria del usuario. " +
                                      "Para re-habilitar, reinstale la aplicación (el instalador elimina este archivo).");
                    return;
                }

                // ================================================================
                // PASO 0.5: Obtener vault activo (sin mutar estado global todavía)
                // ================================================================
                string preLoadedVault = GetActiveVaultPath();

                LoggerHelper.Info("[SyncContext] Iniciando sistema de sync...");

                // Verificar disponibilidad de CR-SQLite (ya inicializado en AutoOpen)
                if (!CrSqliteLoader.IsAvailable)
                {
                    LoggerHelper.Warning("[SyncContext] CR-SQLite no disponible. Sync deshabilitado.");
                    return;
                }

                // Abrir una conexión temporal para consultar la subunidad del usuario
                bool vaultChanged = false;
                string resolvedSubunidad = ResolveSubunidadWithDialog(preLoadedVault, out vaultChanged);
                if (string.IsNullOrEmpty(resolvedSubunidad))
                {
                    // El usuario canceló el diálogo o no se pudo determinar la subunidad
                    LoggerHelper.Warning("[SyncContext] Subunidad no resuelta. Sync deshabilitado.");
                    return;
                }

                ActiveSubunidad = resolvedSubunidad;

                // FASE 1 FIX / FASES 5-7: Inicializar el contexto definitivo solo si el vault cambió
                // Esto evita el doble Dispose() de la conexión compartida en un arranque normal (Fase 6)
                // y previene la doble inicialización de CRDTs y re-encriptaciones innecesarias (Fase 5).
                if (vaultChanged)
                {
                    LoggerHelper.Info($"[SyncContext] El vault cambió a {resolvedSubunidad}. Ejecutando migraciones y CRDT...");
                    SqliteHelper.ResetState(); // Limpiar caché de conexiones
                    SqliteHelper.GetContext(); // Ejecutar migraciones o Seed (variable newCtx removida - Fase 7)
                    SyncInitializer.EnableAllCrdts(); // Configurar CRDT
                }

                // Construir las dependencias del sistema de sync de forma asíncrona (no bloqueante)
                taskLaunched = true;
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        var pathResolver = new SyncPathResolver();
                        string vaultPath = SqliteHelper.GetVaultPath();

                        var syncService = new SubunidadSyncService(
                            subunidad: resolvedSubunidad,
                            vaultPath: vaultPath,
                            pathResolver: pathResolver);

                        Scheduler = new SyncScheduler(syncService);

                        _watcher = new SyncFileWatcher(syncService, pathResolver, resolvedSubunidad);
                        _watcher.Start();

                        IsEnabled = true;

                        LoggerHelper.Info(
                            $"[SyncContext] Sistema de sync inicializado exitosamente en 2do plano. " +
                            $"Subunidad: {ActiveSubunidad}");
                    }
                    catch (Exception exTask)
                    {
                        LoggerHelper.Error("[SyncContext] Error asíncrono iniciando Watchers.", exTask);
                        IsEnabled = false;
                    }
                    finally 
                    {
                        // FASE 3 FIX: Marcar el estado como finalizado solo cuando termine el hilo
                        System.Threading.Interlocked.Exchange(ref _syncInitState, 2);
                    }
                });

            }
            catch (Exception ex)
            {
                // El sync nunca debe bloquear el arranque del add-in
                LoggerHelper.Error("[SyncContext] Error inicializando el sistema de sync. " +
                                   "El add-in funcionará sin sincronización.", ex);
                IsEnabled = false;
            }
            finally
            {
                // Solo marcamos como completado aquí si hubo un error o un early return ANTES de lanzar el hilo.
                if (!taskLaunched)
                {
                    System.Threading.Interlocked.Exchange(ref _syncInitState, 2);
                }
            }
        }

        /// <summary>
        /// Fuerza un escaneo y procesamiento manual (PULL) de todos los archivos .crsync 
        /// pendientes en los directorios, invocable desde la interfaz de usuario.
        /// </summary>
        public static async System.Threading.Tasks.Task ForcePullAsync()
        {
            if (!IsEnabled || _watcher == null)
            {
                LoggerHelper.Warning("[SyncContext] Intento de ForcePullAsync pero Sync no está habilitado o _watcher es null.");
                return;
            }

            await _watcher.ForceProcessPendingFilesAsync();
        }

        /// <summary>
        /// Detiene el sistema de sync y libera todos los recursos.
        /// Debe llamarse desde ExcelAddIn.AutoClose().
        /// </summary>
        internal static void Shutdown()
        {
            try
            {
                LoggerHelper.Info("[SyncContext] Deteniendo sistema de sync...");

                _watcher?.Stop();
                _watcher?.Dispose();
                _watcher = null;

                Scheduler?.Dispose();
                Scheduler = null;

                IsEnabled = false;
                ActiveSubunidad = null;

                LoggerHelper.Info("[SyncContext] Sistema de sync detenido.");
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("[SyncContext] Error deteniendo el sistema de sync.", ex);
            }
        }

        internal static string GetActiveVaultPath()
        {
            try
            {
                string activeVaultPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ConversorPDF", "Data", ".active_vault");

                if (!File.Exists(activeVaultPath))
                    return null;

                string savedPath = File.ReadAllText(activeVaultPath).Trim();
                if (string.IsNullOrWhiteSpace(savedPath))
                    return null;

                if (!File.Exists(savedPath))
                {
                    LoggerHelper.Warning($"[SyncContext] active_vault apunta a un archivo inexistente: {savedPath}. Ignorando.");
                    return null;
                }

                return savedPath;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncContext] No se pudo cargar el active_vault: {ex.Message}");
                return null;
            }
        }

        // ====================================================================
        // RESOLUCIÓN DE SUBUNIDAD CON DIÁLOGO
        // ====================================================================

        /// <summary>
        /// Intenta resolver la subunidad del usuario actual.
        /// Si no está registrado, muestra el diálogo bloqueante FormRegistroSubunidad.
        ///
        /// FLUJO CORREGIDO (Arquitectura por Subunidad):
        ///   1. Consulta la subunidad en la DB base (`AcademicData.db.enc`).
        ///   2. Si no existe, muestra el diálogo para que el usuario elija su subunidad.
        ///   3. Al confirmar, SE COPIA la DB base a un nuevo archivo específico de la subunidad
        ///      (ej: `AcademicData_VALLE.db.enc`).
        ///   4. Enruta el resto de la sesión de Excel a ese nuevo archivo usando SetTestVaultPath.
        /// </summary>
        private static string ResolveSubunidadWithDialog(string preLoadedVault, out bool vaultChanged)
        {
            var resolver = new SubunidadResolver();
            string subunidad = null;
            bool newlyRegistered = false;
            string nombre = null;
            string windowsUser = null;
            vaultChanged = false;

            // FASE 9 FIX: Si tenemos un preLoadedVault (el active_vault del usuario), lo asignamos 
            // ANTES de abrir el contexto, pero SOLO dentro del pipeline de resolución, 
            // eliminando el riesgo de re-entrada.
            if (!SqliteHelper.IsTestVaultPathActive && !string.IsNullOrEmpty(preLoadedVault) && !string.Equals(SqliteHelper.GetVaultPath(), preLoadedVault, StringComparison.OrdinalIgnoreCase))
            {
                SqliteHelper.SetTestVaultPath(preLoadedVault);
                LoggerHelper.Info($"[SyncContext] Pre-cargando vault activo: {preLoadedVault}");
            }

            // Usar el contexto base existente para consultar
            using (var context = SqliteHelper.GetContextForPath())
            {
                IDbConnection connection = context.Open();
                context.MarkAsReadOnly();

                // Intentar resolver la subunidad directamente en el vault actual
                var result = resolver.Resolve(connection);
                if (result.IsSuccess)
                {
                    subunidad = result.Value;
                }
                else
                {
                    // El usuario no está registrado. Mostrar diálogo bloqueante.
                    LoggerHelper.Info(
                        $"[SyncContext] Usuario no registrado: {result.Error}. " +
                        $"Mostrando diálogo de registro...");

                    using (var dialog = new FormRegistroSubunidad())
                    {
                        DialogResult dialogResult = dialog.ShowDialog();

                        if (dialogResult != DialogResult.OK)
                        {
                            // El usuario canceló. Sync deshabilitado.
                            LoggerHelper.Info("[SyncContext] Registro cancelado por el usuario. " +
                                               "Creando archivo de exclusión permanente.");
                            CreateSyncOptOutFile();
                            return null;
                        }

                        // Guardar datos temporalmente para registro posterior
                        subunidad = dialog.SubunidadSeleccionada;
                        nombre = dialog.NombreCompleto;
                        windowsUser = Environment.UserName;

                        newlyRegistered = true;
                    }
                }
            }

            // AHORA: Enrutar al vault específico por subunidad
            string newVaultPath = SqliteHelper.GetDbPathForSubunidad(subunidad);
            string currentVaultPath = SqliteHelper.GetVaultPath();

            if (!System.IO.File.Exists(newVaultPath))
            {
                // SOLUCIÓN AL BUG: Se debe copiar la BD limpia original, no la BD activa (que puede ser de otra subunidad)
                string trueBaseVaultPath = System.IO.Path.Combine(SqliteHelper.DbFolder, "AcademicData.db.enc");

                if (System.IO.File.Exists(trueBaseVaultPath))
                {
                    System.IO.File.Copy(trueBaseVaultPath, newVaultPath);
                    LoggerHelper.Info($"[SyncContext] Vault subunidad creado a partir de la DB base limpia: {newVaultPath}");

                    // FASES 64, 65, 66, 69 FIX: Reutilizar un único contexto (resetContext) para verificar integridad del esquema
                    // utilizando el conjunto completo de 21 tablas base (BASE_TABLES) y registrando de forma precisa si
                    // la base requirió reparación en la apertura previa.
                    using (var resetContext = SqliteHelper.GetContextForPath(newVaultPath))
                    {
                        IDbConnection resetConn = resetContext.Open();

                        bool reparado = SqliteHelper.WasSchemaMigratedForPath(newVaultPath);
                        int tablasBase = resetConn.QueryFirstOrDefault<int>(
                            "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name IN (" +
                            string.Join(",", SchemaMigrator.BASE_TABLES.Select(t => $"'{t}'")) + ");");

                        if (tablasBase < SchemaMigrator.BASE_TABLES.Length)
                        {
                            LoggerHelper.Error($"[SyncContext] Vault copiado incompleto tras reparación ({tablasBase}/{SchemaMigrator.BASE_TABLES.Length} tablas base). Verificar origen 'AcademicData.db.enc'.");
                        }
                        else if (reparado)
                        {
                            LoggerHelper.Warning($"[SyncContext] Vault copiado requirió reparación de esquema (Fases 58-62, 65). Origen 'AcademicData.db.enc' posiblemente inconsistente.");
                        }
                        else
                        {
                            LoggerHelper.Info($"[SyncContext] Vault copiado verificado: {tablasBase}/{SchemaMigrator.BASE_TABLES.Length} tablas base.");
                        }

                        string deviceId = SubunidadSyncService.GetCompositeDeviceId(newVaultPath);

                        // CRITICO: Asegurar que la columna last_pushed_db_version exista antes de intentar resetearla
                        new SyncControlRepository().EnsureTablesExist(resetConn);

                        resetConn.Execute(@"
                            INSERT INTO sync_control (subunidad, device_id, last_pushed_db_version, last_pushed_seq, last_pulled_at)
                            VALUES (@Sub, @Dev, 0, 0, NULL)
                            ON CONFLICT(subunidad, device_id) DO UPDATE SET 
                                last_pushed_db_version = 0,
                                last_pushed_seq = 0;",
                            new { Sub = subunidad, Dev = deviceId });
                        resetContext.CloseConnectionResiliently();
                        byte[] resetBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(resetContext.TempDbPath);
                        if (resetBytes != null)
                        {
                            SqliteHelper.UpdateVaultDirectly(resetBytes, newVaultPath);
                            resetContext.MarkChangesPersisted();
                        }
                    }
                }
                else
                {
                    LoggerHelper.Warning($"[SyncContext] Vault base limpio no existe. El nuevo vault se inicializará vacío: {newVaultPath}");
                }
                
                if (!SqliteHelper.IsTestVaultPathActive)
                {
                    vaultChanged = true;
                    SqliteHelper.SetTestVaultPath(newVaultPath);
                }
            }
            else if (!SqliteHelper.IsTestVaultPathActive && !string.Equals(currentVaultPath, newVaultPath, StringComparison.OrdinalIgnoreCase))
            {
                // Si el usuario ya existía y las rutas no coinciden, redirigir
                vaultChanged = true;
                SqliteHelper.SetTestVaultPath(newVaultPath);
            }

            // Registrar al usuario O actualizar configuración si el vault cambió
            if (newlyRegistered || vaultChanged)
            {
                using (var newContext = SqliteHelper.GetContextForPath(newVaultPath))
                {
                    IDbConnection newConn = newContext.Open();
                    
                    if (newlyRegistered)
                    {
                        // Asegurar que las tablas existen (migración V24 debería haberlas creado)
                        var controlRepo = new SyncControlRepository();
                        controlRepo.EnsureTablesExist(newConn);
                        
                        const string insertSql = @"
                            INSERT OR IGNORE INTO usuarios (nombre, nombreusuario, subunidad, activo)
                            VALUES (@Nombre, @NombreUsuario, @Subunidad, 1);";
                        
                        newConn.Execute(insertSql, new
                        {
                            Nombre = nombre,
                            NombreUsuario = windowsUser.ToLowerInvariant(),
                            Subunidad = subunidad.ToUpperInvariant()
                        });
                    }
                    
                    // CRITICO: Asegurar que la configuración global refleje la subunidad actual
                    // Esto soluciona el problema donde la UI abre mostrando una territorial incorrecta
                    try
                    {
                        newConn.Execute("UPDATE config SET subunidad = @Subunidad WHERE id = 1;", new { Subunidad = subunidad.ToUpperInvariant() });
                    }
                    catch (Exception ex)
                    {
                        LoggerHelper.Warning($"[SyncContext] No se pudo actualizar config.subunidad: {ex.Message}");
                    }
                    
                    // Persistir los cambios al vault de subunidad
                    newContext.CloseConnectionResiliently();
                    byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(newContext.TempDbPath);
                    if (plainBytes != null)
                    {
                        SqliteHelper.UpdateVaultDirectly(plainBytes, newVaultPath);
                        newContext.MarkChangesPersisted();
                    }
                }
                
                if (newlyRegistered)
                {
                    LoggerHelper.Info($"[SyncContext] Usuario registrado en vault de subunidad: {newVaultPath}");
                }
                else
                {
                    LoggerHelper.Info($"[SyncContext] Configuración actualizada para la nueva subunidad: {subunidad}");
                }
            }

            // Persistir el active vault para que sobreviva a la recarga del Add-in
            try
            {
                string activeVaultPath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ConversorPDF", "Data", ".active_vault");
                System.IO.File.WriteAllText(activeVaultPath, newVaultPath);
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncContext] No se pudo guardar el active_vault: {ex.Message}");
            }

            // NOTA: No se llama a SqliteHelper.GetContext().Dispose() de forma agresiva,
            // ya que se asume que las capas superiores reaccionarán correctamente
            // y la caché de conexiones manejará el cambio de vault.

            return subunidad;
        }

        // ====================================================================
        // OPT-OUT HELPERS
        // ====================================================================

        /// <summary>
        /// Retorna la ruta completa del archivo sentinela .sync_optout.
        /// </summary>
        private static string GetSyncOptOutPath()
        {
            string dbFolder = Path.GetDirectoryName(SqliteHelper.GetVaultPath())
                ?? Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ConversorPDF", "Data");
            
            return Path.Combine(dbFolder, SyncOptOutFileName);
        }

        /// <summary>
        /// Verifica si el usuario ha optado por no participar en la sincronización.
        /// </summary>
        private static bool IsSyncOptedOut()
        {
            return File.Exists(GetSyncOptOutPath());
        }

        /// <summary>
        /// Crea el archivo sentinela que marca la exclusión voluntaria de sync.
        /// </summary>
        private static void CreateSyncOptOutFile()
        {
            try
            {
                string optOutPath = GetSyncOptOutPath();
                string folder = Path.GetDirectoryName(optOutPath);
                
                if (!string.IsNullOrEmpty(folder) && !Directory.Exists(folder))
                    Directory.CreateDirectory(folder);
                
                File.WriteAllText(optOutPath,
                    $"# Sync Exclusion File\n" +
                    $"# Creado: {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n" +
                    $"# Usuario: {Environment.UserName}\n" +
                    $"# PC: {Environment.MachineName}\n" +
                    $"#\n" +
                    $"# El usuario optó por no participar en la sincronización.\n" +
                    $"# Para re-habilitar, reinstale la aplicación o borre este archivo.");
                
                LoggerHelper.Info($"[SyncContext] Archivo de exclusión creado: {optOutPath}");
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncContext] No se pudo crear archivo de exclusión: {ex.Message}");
            }
        }
    }
}
