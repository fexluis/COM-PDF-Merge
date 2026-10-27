using System;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// Inicializa las tablas CRDT después de que todas las migraciones estén completas.
    /// Se ejecuta FUERA de la transacción del SchemaMigrator porque crsql_as_crr()
    /// no es compatible con transacciones externas.
    /// </summary>
    internal static class SyncInitializer
    {
        // Tablas que serán CRDT
        private static readonly string[] CrdtTables = 
        {
            "terceros", "sedes", "programas", "asignaturas", 
            "microcurriculo", "contenidos", "festivos", "periodos",
            "vinculaciones", "sesiones", "salvedades_conflictos",
            "BaseGastosTerrestre", "config", "usuarios", "terminales",
            "cdpDependencias", "permisos", "terceros_asignaturas"
        };

        private static bool _crdtsEnabled = false;
        private static readonly Dictionary<string, HashSet<string>> _failedCrdtTablesByVault = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        private static readonly Dictionary<string, HashSet<string>> _crdtEnabledTablesByVault = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        private static HashSet<string> GetFailedCrdtTables(string vaultPath)
        {
            string key = vaultPath ?? "default";
            if (!_failedCrdtTablesByVault.TryGetValue(key, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _failedCrdtTablesByVault[key] = set;
            }
            return set;
        }

        private static HashSet<string> GetCrdtEnabledTables(string vaultPath)
        {
            string key = vaultPath ?? "default";
            if (!_crdtEnabledTablesByVault.TryGetValue(key, out var set))
            {
                set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                _crdtEnabledTablesByVault[key] = set;
            }
            return set;
        }

        private static readonly HashSet<string> _secondaryVaultsEnabled = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Fase 25: Candado de exclusión mutua para proteger las colecciones estáticas
        // frente a llamadas concurrentes desde los hilos de sincronización en segundo plano.
        private static readonly object _syncLock = new object();

        internal static void ResetStateForTests()
        {
            lock (_syncLock)
            {
                _crdtsEnabled = false;
                _failedCrdtTablesByVault.Clear();
                _crdtEnabledTablesByVault.Clear();
                _secondaryVaultsEnabled.Clear();
            }
        }

        /// <summary>
        /// Convierte todas las tablas a CRDT (idempotente y thread-safe).
        /// Fase 25: Usa double-check locking para garantizar que la inicialización
        /// ocurra exactamente una vez, incluso cuando múltiples hilos de sincronización
        /// llaman a este método concurrentemente.
        /// Fase 42: Soporte para vault opcional.
        /// FIX-LOG1-LOG2: El aislamiento del caché `failedTables` y `enabledTables` por 
        /// vault ahora se maneja internamente para evitar cruce de estados entre subunidades.
        /// </summary>
        /// <param name="targetVaultPath">Ruta al vault secundario (null para el maestro)</param>
        internal static void EnableAllCrdts(string targetVaultPath = null)
        {
            // Fase 25: Early-return rápido sin lock (optimización de ruta caliente)
            // Fase 49: Short-circuit para sub-vaults usando hashset dedicado
            if (string.IsNullOrEmpty(targetVaultPath))
            {
                if (_crdtsEnabled) return;
            }
            else
            {
                lock (_syncLock)
                {
                    if (_secondaryVaultsEnabled.Contains(targetVaultPath)) return;
                }
            }

            lock (_syncLock)
            {
                // Fase 25: Double-check dentro del lock
                if (string.IsNullOrEmpty(targetVaultPath))
                {
                    if (_crdtsEnabled) return;
                }
                else
                {
                    // Fase 53: Double-check para vaults secundarios dentro del lock principal
                    if (_secondaryVaultsEnabled.Contains(targetVaultPath)) return;
                }

                LoggerHelper.Info("[SyncInitializer] Iniciando conversión de tablas a CRDT...");

                var swStep = System.Diagnostics.Stopwatch.StartNew();

                using (var context = SqliteHelper.GetContextForPath(targetVaultPath))
                {
                    IDbConnection connection = context.Open();

                    var sqliteConn = connection as SQLiteConnection;
                    if (sqliteConn == null || !CrSqliteLoader.EnableOnConnection(sqliteConn))
                    {
                        LoggerHelper.Warning("[SyncInitializer] CR-SQLite no disponible. No se pueden habilitar CRDTs.");
                        context.MarkAsReadOnly(); // Fase 26: Evitar warning de vault sin persistir
                        return;
                    }

                    bool anyChanges = false;
                    string vaultPath = context.EncryptedFilePath;
                    bool isTargetVaultNull = string.IsNullOrEmpty(targetVaultPath);
                    bool schemaMigrated = SqliteHelper.WasSchemaMigratedForPath(vaultPath);

                    bool allSucceeded = true;
                    if (schemaMigrated)
                    {
                        if (isTargetVaultNull)
                        {
                            _crdtsEnabled = false; // Forzar re-ejecución local
                        }
                        else
                        {
                            _secondaryVaultsEnabled.Remove(targetVaultPath);
                        }

                        if (isTargetVaultNull)
                        {
                            _crdtEnabledTablesByVault.Remove("default");
                            _failedCrdtTablesByVault.Remove("default");
                        }
                        else
                        {
                            _crdtEnabledTablesByVault.Remove(targetVaultPath);
                            _failedCrdtTablesByVault.Remove(targetVaultPath);
                        }
                        LoggerHelper.Info($"[SyncInitializer] Esquema migrado detectado para '{vaultPath}'. Regenerando triggers.");
                    }
                    else
                    {
                        // Bloque else (Fase 18) para indicar que procesamos normal
                        LoggerHelper.Debug($"[SyncInitializer] Esquema sin migración pendiente para '{vaultPath}'. Se procesarán tablas CRDT faltantes.");
                    }

                    // Deduplicación idempotente post-merge antes de evaluar y autosanar triggers CRDT
                    int eliminadosUsuarios = DeduplicateUsuarios(connection);
                    if (eliminadosUsuarios > 0) anyChanges = true;

                    var failedTables = GetFailedCrdtTables(targetVaultPath);
                    var enabledTables = GetCrdtEnabledTables(targetVaultPath);

                    foreach (string tableName in CrdtTables)
                    {
                        if (failedTables.Contains(tableName))
                        {
                            LoggerHelper.Debug($"[SyncInitializer] '{tableName}' falló previamente. Omitiendo.");
                            continue;
                        }

                        if (enabledTables.Contains(tableName))
                        {
                            continue; // Cache: Ya se habilitó exitosamente en esta sesión.
                        }

                        bool changed = false;
                        bool success = EnableSingleCrdt(connection, tableName, schemaMigrated, out changed);

                        if (!success)
                        {
                            allSucceeded = false;
                            failedTables.Add(tableName);
                        }
                        else
                        {
                            enabledTables.Add(tableName);
                            if (changed) anyChanges = true;
                        }
                    }

                    // Limpiar columnas fantasma en TODAS las tablas de reloj (ej. post-migración o recreación)
                    // Esto evita el fatal "SQL logic error" al consultar crsql_changes.
                    CleanGhostColumnsInClockTables(connection, CrdtTables);

                    if (allSucceeded)
                    {
                        LoggerHelper.Info($"[SyncInitializer] Conversión CRDT completada exitosamente para '{vaultPath}'.");
                        if (isTargetVaultNull) 
                        {
                            _crdtsEnabled = true;
                        }
                        else
                        {
                            _secondaryVaultsEnabled.Add(targetVaultPath);
                        }
                    }
                    else
                    {
                        LoggerHelper.Warning($"[SyncInitializer] Conversión CRDT completada con fallos para '{vaultPath}'.");
                    }

                    // Persistir cambios solo si hubo modificaciones reales (crsql_as_crr o migraciones DDL)
                    if (anyChanges || schemaMigrated)
                    {
                        // Fase 34 FIX: Log explicativo cuando se persiste solo por migración de esquema
                        if (!anyChanges && schemaMigrated)
                        {
                            LoggerHelper.Debug("[SyncInitializer] Solo cambios de esquema detectados. Persistiendo...");
                        }

                        context.CloseConnectionResiliently();
                        byte[] plainBytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
                        if (plainBytes != null)
                        {
                            SqliteHelper.UpdateVaultDirectly(plainBytes, vaultPath);
                            context.MarkChangesPersisted();
                        }
                        else
                        {
                            LoggerHelper.Error("[SyncInitializer] FALLO: No se pudieron leer bytes de temp DB. Cambios CRDT NO persistidos.");
                            context.MarkAsReadOnly(); // Fase 27: Suprimir warning espurio de Dispose
                        }
                    }
                    else
                    {
                        context.MarkAsReadOnly();
                        LoggerHelper.Info("[SyncInitializer] No hay cambios CRDT ni migraciones pendientes. Omitiendo re-encriptación y persistencia.");
                    }

                    swStep.Stop();
                    LoggerHelper.Debug($"[SyncInitializer] Conversión CRDT completada en {swStep.ElapsedMilliseconds}ms.");
                } // end using context
            } // end lock (_syncLock)
        }

        private static bool EnableSingleCrdt(IDbConnection connection, string tableName, bool schemaMigrated, out bool changed)
        {
            changed = false;
            try
            {
                if (!TableExists(connection, tableName))
                {
                    LoggerHelper.Debug($"[SyncInitializer] '{tableName}' no existe. Omitiendo.");
                    return true;
                }

                // Obtener la columna PK
                string pkColumn = GetPrimaryKeyColumn(connection, tableName);
                if (string.IsNullOrEmpty(pkColumn))
                {
                    LoggerHelper.Warning($"[SyncInitializer] '{tableName}' no tiene PK. No se puede convertir a CRDT.");
                    return false;
                }

                bool conflictUnique = HasConflictingUniqueIndex(connection, tableName);
                bool isAlreadyCrdt = TableHasCrdtClock(connection, tableName);

                // Paso 2: Hacer la PK NOT NULL (recreando la tabla si es necesario)
                bool rebuilt = EnsurePrimaryKeyNotNull(connection, tableName, pkColumn);
                
                // Fase 36 FIX: Refrescar bandera conflictUnique si EnsurePrimaryKeyNotNull eliminó los índices UNIQUE
                if (rebuilt)
                {
                    conflictUnique = HasConflictingUniqueIndex(connection, tableName);
                }

                // Paso 1: Eliminar índices UNIQUE conflictivos ANTES del short circuit (Fase 12)
                if (tableName == "salvedades_conflictos")
                {
                    connection.Execute("DROP INDEX IF EXISTS idx_salvedades_periodo_tercero;");
                    LoggerHelper.Debug("[SyncInitializer] Índice UNIQUE eliminado de salvedades_conflictos.");
                }

                // DROP del índice UNIQUE antes de crsql_as_crr (solo para usuarios)
                if (tableName == "usuarios")
                {
                    connection.Execute("DROP INDEX IF EXISTS idx_usuarios_nombreusuario;");
                    LoggerHelper.Debug("[SyncInitializer] Índice idx_usuarios_nombreusuario eliminado para permitir CRDT.");
                }

                string safeTableName = tableName.Replace("'", "''");

                // FASE 755 & 758: Si la tabla ya es CRDT pero perdió sus triggers (ej. tras drop defensivo de migración),
                // purgar remanentes parciales y restaurarlos limpiamente vía crsql_as_crr antes de evaluar short-circuit.
                if (isAlreadyCrdt && !rebuilt && !TableHasCrdtTriggers(connection, tableName))
                {
                    connection.Execute($"DROP TRIGGER IF EXISTS [{safeTableName}__crsql_dtrig];");
                    connection.Execute($"DROP TRIGGER IF EXISTS [{safeTableName}__crsql_utrig];");
                    connection.Execute($"DROP TRIGGER IF EXISTS [{safeTableName}__crsql_itrig];");
                    connection.Execute($"SELECT crsql_as_crr('{safeTableName}');");
                    LoggerHelper.Info($"[SyncInitializer] Triggers CRDT restaurados en '{tableName}'.");
                    changed = true;
                    return true;
                }

                // Verificación de short-circuit CRDT
                if (isAlreadyCrdt && !rebuilt && !conflictUnique && !schemaMigrated)
                {
                    LoggerHelper.Debug($"[SyncInitializer] '{tableName}' ya es CRDT y el esquema es válido. Omitiendo crsql_as_crr.");
                    return true;
                }

                changed = true;

                // Paso 3: Convertir a CRDT o regenerar triggers
                
                // FIX (Fase 1): Evaluar `!rebuilt`. Si EnsurePrimaryKeyNotNull reconstruyó la tabla,
                // borró los triggers nativos CRDT, por lo que isAlreadyCrdt es obsoleto y debemos
                // inicializar la tabla con crsql_as_crr en lugar de intentar un crsql_alter_begin.
                if (isAlreadyCrdt && !rebuilt)
                {
                    // Fase 22/23 Fix: crsql_alter_begin/commit deben ejecutarse en secuencia atómica.
                    // Si falla crsql_alter_commit, ejecutamos un rollback explícito del alter
                    // para no dejar la tabla en estado intermedio.
                    try
                    {
                        connection.Execute($"SELECT crsql_alter_begin('{safeTableName}');");
                        connection.Execute($"SELECT crsql_alter_commit('{safeTableName}');");
                        LoggerHelper.Debug($"[SyncInitializer] Triggers CRDT regenerados para '{tableName}' tras migración.");
                    }
                    catch (Exception ex)
                    {
                        // Intentar limpiar el estado intermedio si alter_begin se ejecutó
                        try { connection.Execute($"SELECT crsql_alter_rollback('{safeTableName}');"); }
                        catch { /* ignorar si crsql_alter_rollback no está disponible */ }

                        // FIX (Fase 34): Tolerar incompatibilidad de la DLL embebida que no incluye `crsql_alter_begin`.
                        // Cuando crsql_alter_begin falla con "no such function", los triggers CRDT viejos
                        // quedan atados al esquema antiguo (ej. esperando 17 columnas en lugar de 19 tras ALTER TABLE).
                        // Para evitar crasheos de INSERT/UPDATE ("SQL logic error expected X values, got Y"),
                        // dropeamos los artifacts CRDT manualmente y ejecutamos crsql_as_crr de forma autocurativa.
                        if (ex.Message.IndexOf("no such function", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            LoggerHelper.Debug($"[SyncInitializer] crsql_alter_begin no soportado. Forzando regeneración vía crsql_as_crr para '{tableName}'.");

                            // Limpiar triggers stale para que crsql_as_crr los reinstale frescos con el esquema actual
                            // IMPORTANTE: NO dropear __crsql_clock para no perder el historial de sincronización!
                            string qSafe = safeTableName.Replace("'", "''");
                            connection.Execute($"DROP TRIGGER IF EXISTS [{qSafe}__crsql_dtrig];");
                            connection.Execute($"DROP TRIGGER IF EXISTS [{qSafe}__crsql_utrig];");
                            connection.Execute($"DROP TRIGGER IF EXISTS [{qSafe}__crsql_itrig];");

                            connection.Execute($"SELECT crsql_as_crr('{qSafe}');");
                            LoggerHelper.Debug($"[SyncInitializer] '{tableName}' CRDT regenerado vía crsql_as_crr con esquema actualizado ({schemaMigrated}).");
                        }
                        else
                        {
                            throw; // Re-lanzar para que el catch externo lo registre y marque la tabla como fallida
                        }
                    }
                }
                else
                {
                    connection.Execute($"SELECT crsql_as_crr('{safeTableName}');");
                    LoggerHelper.Debug($"[SyncInitializer] '{tableName}' convertida a CRDT exitosamente.");
                }

                return true;
            }
            catch (Exception ex)
            {
                LoggerHelper.WarningLimitado($"crdtFallo|{tableName}", $"[SyncInitializer] No se pudo convertir '{tableName}' a CRDT: {ex.Message}", maxVeces: 1);
                return false;
            }
        }

        private static string GetPrimaryKeyColumn(IDbConnection connection, string tableName)
        {
            string escaped = tableName.Replace("'", "''");
            var columns = connection.Query(
                $"SELECT name FROM pragma_table_info('{escaped}') WHERE pk > 0 ORDER BY pk;");
            return columns?.FirstOrDefault()?.name?.ToString();
        }

        private static bool TableExists(IDbConnection connection, string tableName)
        {
            int count = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@TableName;",
                new { TableName = tableName });
            return count > 0;
        }

        private static bool TableHasCrdtClock(IDbConnection connection, string tableName)
        {
            string clockTableName = $"{tableName}__crsql_clock";
            int count = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@TableName;",
                new { TableName = clockTableName });
            return count > 0;
        }

        /// <summary>
        /// FASE 755: Verifica si los tres triggers CRDT (insert, update, delete) existen en la base de datos.
        /// </summary>
        private static bool TableHasCrdtTriggers(IDbConnection connection, string tableName)
        {
            string t = tableName.Replace("'", "''");
            return connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='trigger' AND name IN (@i,@u,@d);",
                new { i = $"{t}__crsql_itrig", u = $"{t}__crsql_utrig", d = $"{t}__crsql_dtrig" }) == 3;
        }

        private static bool HasConflictingUniqueIndex(IDbConnection connection, string tableName)
        {
            if (tableName == "BaseGastosTerrestre")
            {
                int count = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='sqlite_autoindex_BaseGastosTerrestre_1';");
                return count > 0;
            }
            if (tableName == "salvedades_conflictos")
            {
                int count = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='idx_salvedades_periodo_tercero';");
                return count > 0;
            }
            if (tableName == "usuarios")
            {
                int count = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='idx_usuarios_nombreusuario';");
                return count > 0;
            }
            return false;
        }

        // ====================================================================
        // SOLUCIÓN COMPLETA para EnsurePrimaryKeyNotNull() (Fase 7)
        // ====================================================================

        private static Dictionary<string, string> DropAllProblematicViews(IDbConnection connection)
        {
            var viewDefinitions = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            
            // Obtener todas las vistas que dependen de tablas sincronizables
            var views = connection.Query<(string name, string sql)>(@"
                SELECT name, sql FROM sqlite_master 
                WHERE type = 'view' 
                  AND (sql LIKE '%terceros%' OR sql LIKE '%sedes%' 
                   OR sql LIKE '%vinculaciones%' OR sql LIKE '%sesiones%'
                   OR sql LIKE '%programas%' OR sql LIKE '%asignaturas%'
                   OR sql LIKE '%terminales%');");
            
            foreach (var view in views)
            {
                viewDefinitions[view.name] = view.sql;
                connection.Execute($"DROP VIEW IF EXISTS [{view.name}];");
                LoggerHelper.Debug($"[SyncInitializer] Vista '{view.name}' dropeada temporalmente.");
            }
            
            return viewDefinitions;
        }

        private static bool EnsurePrimaryKeyNotNull(IDbConnection connection, string tableName, string pkColumn)
        {
            // Obtener DDL original para evaluar si necesita recreación
            string createSql = connection.QueryFirstOrDefault<string>(
                $"SELECT sql FROM sqlite_master WHERE type='table' AND name=@TableName;",
                new { TableName = tableName });

            if (string.IsNullOrEmpty(createSql)) return false; // No reconstruyó

            // Verificar si la PK ya es NOT NULL
            string escaped = tableName.Replace("'", "''");
            var colInfo = connection.QueryFirstOrDefault(
                $"SELECT \"notnull\", \"type\" FROM pragma_table_info('{escaped}') WHERE name = @Pk;",
                new { Pk = pkColumn });

            long notnullValue = colInfo != null ? Convert.ToInt64(colInfo.notnull) : 0;
            
            bool hasFks = createSql.IndexOf("REFERENCES", StringComparison.OrdinalIgnoreCase) >= 0;
            bool hasUnique = createSql.IndexOf("UNIQUE", StringComparison.OrdinalIgnoreCase) >= 0;
            bool needsUniqueRemoval = hasUnique && (tableName == "config" || tableName == "usuarios" || tableName == "salvedades_conflictos" || tableName == "BaseGastosTerrestre");
            
            bool hasNotNullInNonPkColumns = CheckNotNullWithoutDefault(pkColumn, connection, tableName);
            
            // CR-SQLite exige PK explícitamente NOT NULL, y rechaza FKs/UNIQUEs adicionales.
            // Si la tabla posee un campo NOT NULL carente de valor DEFAULT, no puede pasarse por alto.
            if (notnullValue == 1 && !hasFks && !needsUniqueRemoval && !hasNotNullInNonPkColumns)
            {
                LoggerHelper.Debug($"[SyncInitializer] '{tableName}' ya cumple requisitos CRDT (NOT NULL, sin FKs, sin UNIQUE conflictivo). Omitiendo.");
                return false; // No reconstruyó
            }

            // ================================================================
            // DROP DE VISTAS TEMPORAL (si la tabla tiene vistas dependientes)
            // ================================================================
            // FIX (Fase 2): Incluir 'terminales' en el arreglo para que las vistas 
            // que dependen de ella (vw_SesionesDetalle) se eliminen antes de recrear la tabla.
            string[] tablesWithViews = { "terceros", "sedes", "programas", "asignaturas", "vinculaciones", "sesiones", "periodos", "terminales" };
            bool hasViews = tablesWithViews.Contains(tableName.ToLowerInvariant());
            var savedViews = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            
            if (hasViews)
            {
                LoggerHelper.Debug($"[SyncInitializer] '{tableName}' tiene vistas. Dropeando vistas temporalmente...");
                savedViews = DropAllProblematicViews(connection);
            }

            // ================================================================
            // RECREAR TABLA CON PK NOT NULL (sin AUTOINCREMENT)
            // ================================================================

            // Normalizar y modificar SQL (eliminar AUTOINCREMENT, agregar NOT NULL, eliminar FKs)
            string normalizedSql = NormalizeSql(createSql);
            string modifiedSql = ReplacePrimaryKeyWithNotNull(normalizedSql, pkColumn);
            modifiedSql = RemoveForeignKeyConstraints(modifiedSql);
            
            // Eliminar UNIQUE de config y usuarios y BaseGastosTerrestre
            if (tableName == "config" || tableName == "usuarios" || tableName == "BaseGastosTerrestre")
                modifiedSql = RemoveUniqueConstraints(modifiedSql);

            // Recrear tabla
            string tempName = $"_{tableName}_crdt_migration";
            
            // Eliminar UNIQUE de salvedades_conflictos al recrear
            if (tableName == "salvedades_conflictos")
            {
                // Eliminar UNIQUE del DDL (puede estar como tabla o columna)
                modifiedSql = Regex.Replace(modifiedSql, @",\s*UNIQUE\s*\([^)]+\)", "", RegexOptions.IgnoreCase);
            }

            // Eliminar NOT NULL de columnas que no sean la PK primaria (requerido por CR-SQLite si no hay DEFAULT)
            modifiedSql = RemoveNotNullFromNonPkColumns(modifiedSql);

            // Limpiar tabla temporal de migración si quedó residual de un intento previo fallido
            connection.Execute($"DROP TABLE IF EXISTS [{tempName}];");

            using (var tx = connection.BeginTransaction())
            {
                try
                {
                    // Deshabilitar FK constraints antes de modificar esquema dentro de la transacción
                    connection.Execute("PRAGMA foreign_keys = OFF;", transaction: tx);

                    // Reemplazo seguro mediante expresión regular para abarcar nombres sin delimitadores
                    string safeTableName = Regex.Escape(tableName);
                    string replacedSql = Regex.Replace(modifiedSql,
                        $@"\b{safeTableName}\b",
                        tempName,
                        RegexOptions.IgnoreCase);

                    string truncatedSql = replacedSql.Length > 200 
                        ? replacedSql.Substring(0, 200) + "..." 
                        : replacedSql;
                    LoggerHelper.Debug($"[SyncInitializer] SQL generado para '{tableName}': {truncatedSql}");

                    connection.Execute(replacedSql, transaction: tx);
                    // [FIX] Fase 41: Eliminada la deduplicación MIN(id) por nombreusuario.
                    // Generaba pérdida de datos de usuarios (mantenía el ID alfabéticamente menor que 
                    // frecuentemente era la semilla vacía) y causaba resurrección zombie sin tombstones.
                    connection.Execute($"INSERT INTO [{tempName}] SELECT * FROM [{tableName}];", transaction: tx);
                    
                    // Migrar índices
                    MigrateIndices(connection, tx, tableName, tempName);
                    
                    connection.Execute($"DROP TABLE [{tableName}];", transaction: tx);
                    connection.Execute($"ALTER TABLE [{tempName}] RENAME TO [{tableName}];", transaction: tx);
                    
                    tx.Commit();
                    
                    LoggerHelper.Debug($"[SyncInitializer] '{tableName}' recreada con PK NOT NULL.");
                }
                catch (Exception ex)
                {
                    tx.Rollback();
                    LoggerHelper.Warning($"[SyncInitializer] Error recreando '{tableName}': {ex.Message}. " +
                        "La tabla conserva su estructura original (PK nullable). " +
                        "La sincronización CRDT no estará disponible para esta tabla.");
                    return false; // Rebuild falló
                }
            }

            // ================================================================
            // RECREAR VISTAS desde las definiciones CAPTURADAS (NO del schema embebido)
            // ================================================================
            if (hasViews && savedViews.Count > 0)
            {
                LoggerHelper.Info($"[SyncInitializer] Recreando vista(s) desde definiciones capturadas...");
                try
                {
                    foreach (var kvp in savedViews)
                    {
                        string viewName = kvp.Key;
                        // Omitir vistas críticas que serán recreadas con su DDL canónico explícitamente abajo
                        if (viewName.Equals("vw_SesionesDetalle", StringComparison.OrdinalIgnoreCase) || 
                            viewName.Equals("vw_VinculacionesDetalle", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        string viewSql = kvp.Value;
                        connection.Execute(viewSql);
                        LoggerHelper.Debug($"[SyncInitializer] Vista '{viewName}' recreada desde caché temporal.");
                    }
                    
                    // FIX FASE 14/Fase 10: Recrear las vistas críticas con su DDL canónico para evitar propagar 
                    // definiciones stale (antiguas) cacheadas en sqlite_master antes de las migraciones recientes.
                    // Ahora se ejecuta dentro de una transacción para evitar estado inconsistente si la segunda falla.
                    using (var tx = connection.BeginTransaction())
                    {
                        try
                        {
                            // Fase 57 FIX: No ejecutar V33 si 'terminales' no existe (evita crash durante instalación inicial/CRDT incompleto)
                            int termExists = connection.QueryFirstOrDefault<int>(
                                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='terminales';",
                                transaction: tx);
                            
                            if (termExists > 0)
                            {
                                new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV15_AddModalidadVinculacion().Up(connection, tx);
                                new ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations.MigrationV33_UpdateSesionesDetalleForTerminales().Up(connection, tx);
                            }
                            else
                            {
                                // Fase 58 FIX: terminales aún no existe → recrear solo vw_VinculacionesDetalle (no depende de terminales)
                                // vw_SesionesDetalle quedará ausente hasta próxima sesión; el repositorio lo detectará y reintentará.
                                LoggerHelper.Warning("[SyncInitializer] Tabla 'terminales' no existe. Recreando solo vw_VinculacionesDetalle. vw_SesionesDetalle se restaurará en próxima sesión.");
                                connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: tx);
                                connection.Execute(@"
                                    CREATE VIEW vw_VinculacionesDetalle AS
                                    SELECT 
                                        v.id AS IdVinculacion,
                                        v.idsede AS IdSede,
                                        v.idtercero AS IdTercero,
                                        pe.idperiodo AS IdPeriodo,
                                        pe.periodo AS PeriodoNombre,
                                        p.modalidad AS ProgramaTipo,
                                        p.abreviatura AS ProgramaAbreviatura,
                                        t.nombre AS TerceroNombre,
                                        CASE 
                                            WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                                            ELSE v.semestre || '-' || v.grupo 
                                        END AS Semestre,
                                        a.nombre AS AsignaturaNombre,
                                        s.ciudad AS SedeCiudad,
                                        NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS SesionFechaInicial,
                                        NULLIF(NULLIF(v.fecha_fin, ''), '--') AS SesionFechaFinal,
                                        v.total_horas AS SesionHoras,
                                        (SELECT CASE WHEN EXISTS (
                                            SELECT 1 FROM sesiones ses WHERE ses.idvinculacion = v.id AND ses.viatica = 1
                                        ) THEN 1 ELSE 0 END) AS SesionViatica,
                                        v.modalidadVinculacion AS Modalidad,
                                        v.subunidad AS Subunidad
                                    FROM vinculaciones v
                                    JOIN programas p ON v.idprograma = p.id
                                    JOIN asignaturas a ON v.idasignatura = a.id
                                    JOIN sedes s ON v.idsede = s.id
                                    LEFT JOIN terceros t ON v.idtercero = t.id
                                    LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;
                                ", transaction: tx);
                            }
                            
                            tx.Commit();
                            LoggerHelper.Info($"[SyncInitializer] Vistas recreadas exitosamente y corregidas a última versión de forma atómica.");
                            SqliteHelper.EnsureAgViews(connection);
                        }
                        catch (Exception innerEx)
                        {
                            tx.Rollback();
                            throw new Exception("Falló la transacción de rescate de vistas.", innerEx);
                        }
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[SyncInitializer] Error recreando vistas: {ex.Message}. " +
                        "Las vistas se recrearán automáticamente en el próximo reinicio del add-in.");
                    // Fase 23 FIX: Relanzar para que EnableSingleCrdt marque la tabla como fallida y no asuma éxito.
                    throw;
                }
            }

            return true; // Retorna true indicando que el esquema fue reconstruido
        }

        private static string NormalizeSql(string sql)
        {
            // Paso 0: Eliminar comentarios SQL de línea (-- ...) ANTES de colapsar líneas
            // Esto previene que -- trunque el resto del DDL al unificar todo en una línea
            string result = Regex.Replace(sql, @"--[^\r\n]*", "");
            
            result = result.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ");
            while (result.IndexOf("  ", StringComparison.Ordinal) >= 0)
            {
                result = result.Replace("  ", " ");
            }
            return result;
        }

        private static string ReplacePrimaryKeyWithNotNull(string sql, string pkColumn)
        {
            // Enfoque simplificado: buscar patrones con IndexOf y reemplazar con Replace literal
            // Fase 29 FIX: Agregar patrones con NOT NULL preexistente para evitar desajuste de columnas
            string[] patterns = {
                $"\"{pkColumn}\" INTEGER PRIMARY KEY NOT NULL AUTOINCREMENT",
                $"{pkColumn} INTEGER PRIMARY KEY NOT NULL AUTOINCREMENT",
                $"\"{pkColumn}\" INTEGER PRIMARY KEY AUTOINCREMENT",
                $"{pkColumn} INTEGER PRIMARY KEY AUTOINCREMENT",
                $"\"{pkColumn}\" INTEGER PRIMARY KEY NOT NULL",
                $"{pkColumn} INTEGER PRIMARY KEY NOT NULL",
                $"\"{pkColumn}\" INTEGER PRIMARY KEY",
                $"{pkColumn} INTEGER PRIMARY KEY",
                $"\"{pkColumn}\" TEXT PRIMARY KEY NOT NULL",
                $"{pkColumn} TEXT PRIMARY KEY NOT NULL",
                $"\"{pkColumn}\" TEXT PRIMARY KEY",
                $"{pkColumn} TEXT PRIMARY KEY"
            };
            string replacementInt = $"{pkColumn} INTEGER PRIMARY KEY NOT NULL";
            string replacementText = $"{pkColumn} TEXT PRIMARY KEY NOT NULL";

            string result = sql;
            foreach (string pattern in patterns)
            {
                int idx = result.IndexOf(pattern, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                {
                    bool isText = pattern.IndexOf(" TEXT ", StringComparison.OrdinalIgnoreCase) >= 0;
                    string replacement = isText ? replacementText : replacementInt;
                    
                    result = result.Substring(0, idx) + replacement + 
                             result.Substring(idx + pattern.Length);
                    break;
                }
            }
            return result; // Si no encuentra, retorna el original
        }

        /// <summary>
        /// Elimina todo tipo de constraints UNIQUE de un statement CREATE TABLE.
        /// FIX: Modificado para soportar UNIQUE a nivel de columna, requisito estricto
        /// de CR-SQLite (ej: BaseGastosTerrestre con `CodigoTrayecto TEXT UNIQUE`).
        /// </summary>
        private static string RemoveUniqueConstraints(string sql)
        {
            // Paso 0: Eliminar constraints UNIQUE a nivel de columna (ej. "CodigoTrayecto TEXT NOT NULL UNIQUE")
            string result = Regex.Replace(sql, @"\s+UNIQUE\b(?!\s*\()", "", RegexOptions.IgnoreCase);

            // Paso 1: Eliminar constraints UNIQUE a nivel de tabla: ", UNIQUE(col1, col2)"
            //         Se captura la coma precedente y el grupo de columnas entre paréntesis.
            //         El patrón maneja tanto si va seguido de coma (intermedio) como de cierre ") " (final).
            result = Regex.Replace(result, @",\s*UNIQUE\s*\([^)]+\)", "", RegexOptions.IgnoreCase);
            
            // Paso 1b: Eliminar CONSTRAINT ... UNIQUE(col)
            result = Regex.Replace(result, @",\s*CONSTRAINT\s+\w+\s+UNIQUE\s*\([^)]+\)", "", RegexOptions.IgnoreCase);
            
            // Paso 1c: Eliminar CONSTRAINT sin coma precedente (si es la última columna)
            result = Regex.Replace(result, @"CONSTRAINT\s+\w+\s+UNIQUE\s*\([^)]+\)\s*,?", "", RegexOptions.IgnoreCase);
            
            // Paso 2: Eliminar constraint UNIQUE a nivel de columna: "TEXT NOT NULL UNIQUE"
            //         Solo donde UNIQUE NO va seguido de paréntesis (para no duplicar con el paso 1).
            result = Regex.Replace(result, @"\s+UNIQUE(?!\s*\()", " ", RegexOptions.IgnoreCase);
            
            // Paso 3: Colapsar espacios múltiples generados por los reemplazos
            result = Regex.Replace(result, @"\s+", " ");
            
            // Paso 4: Limpiar paréntesis con contenido vacío residual ", )" → ")"
            result = Regex.Replace(result, @",\s*\)", ")");
            
            return result;
        }

        private static string RemoveForeignKeyConstraints(string sql)
        {
            // Elimina FOREIGN KEY(col) REFERENCES tabla(col) [ON DELETE/UPDATE ...]
            // Captura la coma precedente si existe, o lo deja limpio si es el último elemento
            return Regex.Replace(sql, 
                @"\s*,?\s*FOREIGN\s+KEY\s*\([^)]+\)\s*REFERENCES\s+\w+\s*\([^)]+\)([^,)]*)", 
                "", RegexOptions.IgnoreCase);
        }

        private static string RemoveNotNullFromNonPkColumns(string sql)
        {
            // Fase 24 FIX: Elimina NOT NULL de columnas que no son la PK.
            // Conserva NOT NULL solo en la columna PK usando enfoque de sustitución de string 
            // (evita ArgumentException por lookbehind de ancho variable en regex .NET 4.8)
            string result = sql.Replace("PRIMARY KEY NOT NULL", "PRIMARY KEY__KEEP__");
            result = Regex.Replace(result, @"\bNOT\s+NULL\b", "", RegexOptions.IgnoreCase);
            result = result.Replace("PRIMARY KEY__KEEP__", "PRIMARY KEY NOT NULL");
            result = Regex.Replace(result, @"\s+", " ");
            return result;
        }

        private static bool CheckNotNullWithoutDefault(string pkColumn, IDbConnection connection, string tableName)
        {
            // Evalúa a través del pragma si existe alguna columna NOT NULL que NO sea la llave primaria primaria y que carezca de un valor por defecto.
            string escaped = tableName.Replace("'", "''");
            var columns = connection.Query(
                $"PRAGMA table_info('{escaped}');");

            foreach (var col in columns)
            {
                string name = (string)col.name;
                long notnull = Convert.ToInt64(col.notnull);
                string dflt_value = (string)col.dflt_value;

                if (name != pkColumn && notnull == 1 && string.IsNullOrEmpty(dflt_value))
                {
                    return true;
                }
            }
            return false;
        }

        private static void MigrateIndices(IDbConnection connection, IDbTransaction tx, string tableName, string tempName)
        {
            var indices = connection.Query(@"
                SELECT il.name AS index_name, il.sql AS index_sql
                FROM sqlite_master il WHERE il.type='index'
                  AND il.tbl_name = @TableName
                  AND il.name NOT LIKE 'sqlite_auto%'
                  AND il.name NOT LIKE '%_crdt_migration';",
                new { TableName = tableName }, transaction: tx);

            foreach (var idx in indices)
            {
                string idxName = (string)idx.index_name;
                string idxSql = (string)idx.index_sql;
                
                // Omitir UNIQUE conflictivos
                if (tableName == "salvedades_conflictos" && idxName == "idx_salvedades_periodo_tercero")
                    continue;
                
                // Omitir idx_usuarios_nombreusuario: se recreará manualmente después del RENAME
                // porque su UNIQUE subyacente colisiona con el DDL modificado de la tabla temporal.
                if (tableName == "usuarios" && idxName == "idx_usuarios_nombreusuario")
                    continue;
                
                if (!string.IsNullOrEmpty(idxSql))
                {
                    // Usar CREATE INDEX IF NOT EXISTS para evitar colisiones
                    string safeSql = Regex.Replace(idxSql, @"\bCREATE\s+(UNIQUE\s+)?INDEX\b", 
                        "CREATE $1INDEX IF NOT EXISTS", RegexOptions.IgnoreCase);
                    connection.Execute(
                        safeSql.Replace($"\"{tableName}\"", $"\"{tempName}\"")
                               .Replace($"[{tableName}]", $"[{tempName}]"),
                        transaction: tx);
                }
            }
        }

        /// <summary>
        /// Elimina filas en __crsql_clock cuyas columnas ya no existen en la tabla base.
        /// Esto previene que la vista virtual `crsql_changes` falle con "SQL logic error" 
        /// (por ej. tras renombrar/eliminar columnas o recrear tablas).
        /// </summary>
        private static void CleanGhostColumnsInClockTables(IDbConnection connection, IEnumerable<string> tables)
        {
            foreach (var table in tables)
            {
                if (!TableExists(connection, table)) continue;

                string clockTable = $"{table}__crsql_clock";
                if (!TableExists(connection, clockTable)) continue;

                try
                {
                    // '-1' es usado internamente por CR-SQLite para registrar inserciones/eliminaciones de filas completas.
                    string sql = $@"
                        DELETE FROM [{clockTable}] 
                        WHERE col_name != '-1' 
                          AND col_name NOT IN (SELECT name FROM pragma_table_info('{table}'));
                    ";
                    
                    int deleted = connection.Execute(sql);
                    if (deleted > 0)
                    {
                        LoggerHelper.Info($"[SyncInitializer] Limpiadas {deleted} columnas fantasma en '{clockTable}' para prevenir 'SQL logic error'.");
                    }
                }
                catch (Exception ex)
                {
                    LoggerHelper.Warning($"[SyncInitializer] Error limpiando '{clockTable}': {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Deduplicación idempotente de usuarios conservando el registro con datos más ricos.
        /// Se ejecuta tras merges/pulls y antes de la autosanación de triggers CRDT.
        /// </summary>
        internal static int DeduplicateUsuarios(IDbConnection connection)
        {
            try
            {
                if (!TableExists(connection, "usuarios"))
                {
                    return 0;
                }

                string sqlDeleteDuplicates = @"
                    DELETE FROM usuarios 
                    WHERE id NOT IN (
                        SELECT id FROM (
                            SELECT id,
                                   ROW_NUMBER() OVER (
                                       PARTITION BY LOWER(TRIM(nombreusuario)) 
                                       ORDER BY 
                                           CASE WHEN firma IS NOT NULL AND firma != '' THEN 1 ELSE 0 END DESC,
                                           CASE WHEN perfil IS NOT NULL AND perfil != 'Academica' THEN 1 ELSE 0 END DESC,
                                           updated_at DESC,
                                           id ASC
                                   ) as rn
                            FROM usuarios
                        )
                        WHERE rn = 1
                    );";

                int rowsDeleted = connection.Execute(sqlDeleteDuplicates);
                if (rowsDeleted > 0)
                {
                    LoggerHelper.Info($"[SyncInitializer] Deduplicación post-merge completada: {rowsDeleted} filas eliminadas en usuarios.");
                }
                return rowsDeleted;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"[SyncInitializer] Error ejecutando deduplicación post-merge de usuarios: {ex.Message}");
                return 0;
            }
        }
    }
}
