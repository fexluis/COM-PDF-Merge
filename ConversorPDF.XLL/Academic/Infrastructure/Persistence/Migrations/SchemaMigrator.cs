using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Dapper;
using ConversorPDF.XLL.Academic.Infrastructure.Persistence.Context;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class SchemaMigrator : ISchemaMigrator
    {
        private readonly IEnumerable<IMigration> _migrations;
        private readonly IVaultPersister _vaultPersister;
        private readonly ILogger _logger;
        
        // Fase 40 y 46 FIX: Cache estático aislado por vault para evitar re-encriptación y race conditions
        private static readonly HashSet<string> _hotfixAppliedVaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _hotfixLock = new object();
        private static readonly Dictionary<int, int> _attempts = new Dictionary<int, int>();

        public SchemaMigrator(IEnumerable<IMigration> migrations, IVaultPersister vaultPersister, ILogger logger)
        {
            _migrations = migrations.OrderBy(m => m.Version);
            _vaultPersister = vaultPersister;
            _logger = logger;
        }

        private const int BASE_SCHEMA_VERSION = 28;

        public static readonly string[] BASE_TABLES =
        {
            "__SchemaVersion", "festivos", "programas", "asignaturas", "microcurriculo", "contenidos",
            "terceros", "territoriales", "sedes", "BaseGastosTerrestre", "periodos", "vinculaciones",
            "sesiones", "salvedades_conflictos", "usuarios", "permisos", "sync_control", "sync_archive",
            "terminales", "config", "cdpDependencias"
        };

        private bool RepairBaseSchemaIfNeeded(IDbConnection connection, string vaultPath = null)
        {
            int declaredVersion = connection.QueryFirstOrDefault<int>(
                "SELECT Version FROM __SchemaVersion ORDER BY Version DESC LIMIT 1;");

            int baseTablesCount = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name IN (" +
                string.Join(",", BASE_TABLES.Select(t => $"'{t}'")) + ");");

            if (declaredVersion >= BASE_SCHEMA_VERSION && baseTablesCount == BASE_TABLES.Length) return false;

            _logger.Warning($"[SchemaMigrator] Esquema base ausente/incompleto (v{declaredVersion}, tablas {baseTablesCount}/{BASE_TABLES.Length}). Reparando...");

            using (var tx = connection.BeginTransaction())
            {
                try
                {
                    string subunidadDestino = SqliteHelper.ExtraerSubunidadDeRuta(vaultPath);
                    SqliteHelper.ApplyBaseSchemaAndSeeds(connection, tx, subunidadDestino);
                    connection.Execute("DELETE FROM __SchemaVersion WHERE Version > @Base;",
                                       new { Base = BASE_SCHEMA_VERSION }, transaction: tx);
                    connection.Execute("INSERT OR IGNORE INTO __SchemaVersion (Version) VALUES (@Base);",
                                       new { Base = BASE_SCHEMA_VERSION }, transaction: tx);
                    tx.Commit();
                }
                catch (Exception ex)
                {
                    try { tx.Rollback(); } catch { }
                    _logger.Error("[SchemaMigrator] Error durante la reparación del esquema base. Transacción revertida.", ex);
                    throw;
                }
            }
            return true;
        }

        public void EnsureSchemaUpToDate(EncryptedSQLiteContext context)
        {
            IDbConnection connection = context.Open();
            
            // 1. Asegurar tabla de versión (Idempotencia)
            connection.Execute("CREATE TABLE IF NOT EXISTS __SchemaVersion (Version INTEGER PRIMARY KEY);");
            
            // 2. FASES 58, 60, 61, 62, 75: Verificar y reparar esquema base antes de calcular migraciones pendientes
            int currentVersion = connection.QueryFirstOrDefault<int>("SELECT Version FROM __SchemaVersion ORDER BY Version DESC LIMIT 1;");
            bool baseSchemaRepaired = RepairBaseSchemaIfNeeded(connection, context.EncryptedFilePath);
            if (baseSchemaRepaired) currentVersion = Math.Min(currentVersion, BASE_SCHEMA_VERSION);

            // FASE 894: Calcular pendientes por versiones no registradas en __SchemaVersion
            // Evita que un hueco por fallo en migración tolerante (ej. V67) sea omitido si se registró una versión superior (V68).
            var aplicadas = new HashSet<int>(
                connection.Query<int>("SELECT Version FROM __SchemaVersion;"));

            // FASE 898: Las versiones <= BASE_SCHEMA_VERSION pertenecen al esquema base (AcademicSchema.sql); nunca re-ejecutar.
            const int baseVersion = BASE_SCHEMA_VERSION;

            var pendingMigrations = _migrations
                .Where(m => !aplicadas.Contains(m.Version) && m.Version > baseVersion)
                .OrderBy(m => m.Version)
                .ToList();

            // Fase 35, 37 y 45 FIX: Auto-trigger MigrationV42 evaluando la ausencia de datos en territoriales
            bool hotfixApplied = false;
            if (currentVersion >= 42)
            {
                lock (_hotfixLock)
                {
                    if (!_hotfixAppliedVaults.Contains(context.EncryptedFilePath))
                    {
                        bool needV42 = false;
                        if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='divipola';") == 0) needV42 = true;
                        if (!needV42 && connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='BaseGastosTerrestre';") == 0) needV42 = true;
                        if (!needV42 && connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='territoriales';") == 0) needV42 = true;
                        if (!needV42 && connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM territoriales;") == 0) needV42 = true;

                        if (needV42)
                        {
                            _logger.Info("[SchemaMigrator] Tablas o datos semilla faltantes. Ejecutando MigrationV42 como hotfix...");
                            var mig42 = _migrations.FirstOrDefault(m => m.Version == 42);
                            if (mig42 != null)
                            {
                                using (var tx = connection.BeginTransaction())
                                {
                                    mig42.Up(connection, tx);
                                    tx.Commit();
                                }
                                hotfixApplied = true;
                                _hotfixAppliedVaults.Add(context.EncryptedFilePath);
                            }
                        }
                    }
                }
            }

            // Fase 30 FIX: Se eliminó needsPersist (residuo de self-healing antiguo)
            if (!pendingMigrations.Any() && !hotfixApplied && !baseSchemaRepaired)
            {
                // No hubo cambios: el vault en disco es la versión canónica.
                // Marcar como persistido para suprimir la advertencia espuria en Dispose().
                _logger.Info($"[SchemaMigrator] Base de datos actualizada (v{currentVersion}). Sin migraciones pendientes.");
                VerificarColumnasRequeridas(connection);
                context.MarkChangesPersisted();
                return;
            }

            foreach (var migration in pendingMigrations)
            {
                _logger.Info($"[SchemaMigrator] Migrando base de datos de v{currentVersion} a v{migration.Version} ({migration.Name})...");
                
                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        migration.Up(connection, transaction);
                        connection.Execute("INSERT OR IGNORE INTO __SchemaVersion (Version) VALUES (@Version);", new { migration.Version }, transaction: transaction);
                        transaction.Commit();
                        
                        currentVersion = migration.Version;
                    }
                    catch (Exception ex)
                    {
                        // Fase 43 FIX: Purgar caché si aborta para permitir reintento de hotfix en próximo arranque
                        if (hotfixApplied) 
                        {
                            lock (_hotfixLock)
                            {
                                _hotfixAppliedVaults.Remove(context.EncryptedFilePath);
                            }
                        }
                        try { transaction.Rollback(); } catch { }

                        // Fase 178, 748 y 867: Migraciones aditivas tolerantes no deben abortar el arranque,
                        // pero se limita y registra el número de intentos para evitar ruido indefinido en logs.
                        if (migration is IMigrationTolerante tolerante && tolerante.Aditiva)
                        {
                            _attempts.TryGetValue(migration.Version, out int nTol);
                            _attempts[migration.Version] = ++nTol;

                            if (nTol >= 3)
                            {
                                _logger.Error($"[SchemaMigrator] v{migration.Version} ({migration.Name}) falló {nTol} veces en esta sesión. " +
                                              "Se omite el resto del arranque; se reintentará en el próximo inicio.");
                            }
                            else
                            {
                                _logger.Error($"[SchemaMigrator] HUECO DE VERSIÓN: v{migration.Version} ({migration.Name}) FALLÓ pero es ADITIVA. Se continúa sin registrar versión: {ex.Message}");
                            }
                            continue;
                        }

                        // Fase 179: Corte de bucle infinito de reintentos
                        _attempts.TryGetValue(migration.Version, out int n);
                        _attempts[migration.Version] = ++n;
                        if (n >= 3)
                        {
                            _logger.Error($"[SchemaMigrator] v{migration.Version} falló {n} veces consecutivas. Se omite para no bloquear el arranque.");
                            continue;
                        }

                        _logger.Error($"[SchemaMigrator] Error ejecutando migración v{migration.Version}", ex);
                        throw;
                    }
                }
            }

            // Fase 197: Autochequeo preventivo de columnas requeridas al finalizar migraciones
            VerificarColumnasRequeridas(connection);

            context.CloseConnectionResiliently();
            byte[] bytes = EncryptedSQLiteContext.LeerBytesConRetrySync(context.TempDbPath);
            if (bytes != null)
            {
                _vaultPersister.Persist(bytes, context.EncryptedFilePath);
                
                // MOVER AQUÍ: Solo marcar si la persistencia fue exitosa
                SqliteHelper.MarkVaultAsMigrated(context.EncryptedFilePath);
                
                _logger.Info($"[SchemaMigrator] Todas las migraciones persistidas exitosamente.");
            }
            else
            {
                _logger.Error($"[SchemaMigrator] FALLO: No se pudieron leer bytes de temp DB en {context.TempDbPath}. Cambios NO persistidos al vault.");
                // Fase 39 FIX: Retornar temprano. No marcar MarkChangesPersisted, permitiendo que Dispose() de EncryptedSQLiteContext 
                // haga un intento de salvataje al apagarse.
                return;
            }
            
            context.MarkChangesPersisted();
        }

        private static readonly Dictionary<string, string[]> ColumnasRequeridas =
            new Dictionary<string, string[]>
        {
            { "periodos",    new[] { "periodo_inicio", "periodo_fin" } },
            { "terceros",    new[] { "cedula", "nombre", "subunidad" } },
            { "asignaturas", new[] { "cod_materia", "int_horaria", "nombre" } },
            { "programas",   new[] { "nombreprograma", "abreviatura", "activo" } },
            { "sedes",       new[] { "ciudad", "activo" } }
        };

        /// <summary>
        /// FASE 197: Inspección preventiva de columnas estructurales críticas tras el proceso de migración.
        /// Registra alertas en log si detecta un esquema incompleto en tablas sensibles sin interrumpir la ejecución.
        /// </summary>
        internal void VerificarColumnasRequeridas(IDbConnection connection)
        {
            try
            {
                foreach (var kv in ColumnasRequeridas)
                {
                    int tableExists = connection.QueryFirstOrDefault<int>(
                        $"SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='{kv.Key}';");
                    if (tableExists == 0) continue;

                    var reales = new HashSet<string>(
                        connection.Query<string>($"SELECT name FROM pragma_table_info('{kv.Key}');")
                                  .Select(x => x.ToLowerInvariant()),
                        StringComparer.OrdinalIgnoreCase);

                    var faltan = kv.Value.Where(c => !reales.Contains(c)).ToArray();
                    if (faltan.Length > 0)
                    {
                        _logger.Error($"[SchemaMigrator] ESQUEMA INCOMPLETO en '{kv.Key}': faltan {string.Join(", ", faltan)}. Acción requerida: aplique la migración de reparación correspondiente.");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning($"[SchemaMigrator] No se pudo verificar integridad de columnas requeridas: {ex.Message}");
            }
        }
    }
}

