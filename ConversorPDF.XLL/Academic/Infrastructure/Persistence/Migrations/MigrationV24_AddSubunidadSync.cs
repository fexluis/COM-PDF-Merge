using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using Dapper;
using ConversorPDF.XLL.Core.Sync;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// MIGRACIÓN V24: Preparar el esquema para sincronización CR-SQLite.
    ///
    /// CAMBIOS QUE APLICA:
    ///   1. Agrega columna `subunidad` a tablas existentes que deben sincronizarse.
    ///   2. Crea las tablas nuevas `config` y `usuarios`.
    ///   3. Crea tablas de control sync_control y sync_archive.
    ///   4. Convierte todas las tablas sincronizables a CRDTs con crsql_as_crr().
    ///   5. Pobla la columna subunidad con valores por defecto para datos históricos.
    ///
    /// IDEMPOTENCIA:
    ///   Cada operación es defensiva: usa PRAGMA table_info para verificar si la columna
    ///   ya existe antes de ALTER TABLE (SQLite no soporta IF NOT EXISTS en ALTER TABLE).
    ///   crsql_as_crr() es idempotente — si la tabla ya es CRDT, no hace nada.
    ///
    /// DEPENDENCIA CRÍTICA:
    ///   crsql_as_crr() requiere la extensión CR-SQLite cargada en la conexión.
    ///   Este método llama CrSqliteLoader.EnableOnConnection() al inicio.
    ///   Si CR-SQLite no está disponible, la migración crea las tablas/columnas
    ///   pero OMITE la conversión a CRDT (el sync no funcionará pero la DB sí).
    ///
    /// DATOS HISTÓRICOS:
    ///   Para registros existentes sin subunidad asignada:
    ///   - terceros/sedes: usa la columna `departamento` si existe.
    ///   - resto de tablas: queda como '' (cadena vacía, no null).
    ///   El administrador puede actualizar manualmente estos valores después.
    /// </summary>
    public class MigrationV24_AddSubunidadSync : IMigration
    {
        // ====================================================================
        // IDENTIFICACIÓN
        // ====================================================================

        public int Version => 24;

        public string Name => "Add Subunidad Column and Enable CR-SQLite CRDTs";

        // ====================================================================
        // TABLAS QUE RECIBEN COLUMNA SUBUNIDAD (existentes)
        // ====================================================================

        /// <summary>
        /// Tablas existentes que necesitan la columna subunidad para el filtrado de sync.
        /// Formato: (nombre de tabla, columna de datos relacionados para retrocompatibilidad).
        /// Si la columna de retrocompatibilidad existe, se usa para poblar subunidad en datos históricos.
        /// </summary>
        private static readonly (string TableName, string FallbackColumn)[] ExistingTablesWithSubunidad =
        {
            ("terceros",              "departamento"),
            ("sedes",                 "departamento"),
            ("programas",             null),
            ("asignaturas",           null),
            ("microcurriculo",        null),
            ("contenidos",            null),
            ("festivos",              null),
            ("periodos",              null),
            ("vinculaciones",         null),
            ("sesiones",              null),
            ("salvedades_conflictos", null),
            ("BaseGastosTerrestre",   null)
        };

        // ====================================================================
        // DDL NUEVAS TABLAS
        // ====================================================================

        /// <summary>
        /// DDL de la tabla config: almacena parámetros de configuración por subunidad.
        /// Cada subunidad tiene exactamente UN registro en esta tabla.
        /// Incluye datos de firmantes, viáticos y ubicación geográfica.
        /// </summary>
        private const string CreateConfigSql = @"
            CREATE TABLE IF NOT EXISTS config (
                id                          INTEGER PRIMARY KEY,
                subunidad                   TEXT NOT NULL,
                id_territorial              INTEGER NOT NULL,
                id_periodo                  INTEGER NOT NULL,
                fecha_inicio                TEXT,
                fecha_fin                   TEXT,
                cdp                         TEXT,
                salario_decreto_613_2025    REAL,
                viatico_diario_tarifa_a     REAL,
                viatico_diario_tarifa_b     REAL,
                porcentaje_decreto_a        REAL,
                porcentaje_decreto_b        REAL,
                dt_nombre                   TEXT,
                dt_cargo                    TEXT,
                corporativo_cedula          TEXT,
                corporativo_nombre          TEXT,
                corporativo_cargo           TEXT,
                departamento_base           TEXT,
                ciudad_base                 TEXT,
                aeropuerto_base             TEXT,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now'))
            );";

        /// <summary>
        /// DDL de la tabla usuarios: mapea cuentas de Windows a subunidades.
        /// La columna nombreusuario coincide con Environment.UserName.
        /// UUID como PK para evitar colisiones en sync multi-dispositivo.
        /// FASE 32: Se eliminaron los comentarios inline (--) del script SQL 
        /// ya que causaban fallos silenciosos de parseo en Dapper/System.Data.SQLite.
        /// </summary>
        private const string CreateUsuariosSql = @"
            CREATE TABLE IF NOT EXISTS usuarios (
                id              TEXT PRIMARY KEY DEFAULT (lower(hex(randomblob(16)))),
                nombre          TEXT NOT NULL,
                nombreusuario   TEXT NOT NULL,
                subunidad       TEXT NOT NULL,
                activo          INTEGER DEFAULT 1,
                created_at TEXT DEFAULT (datetime('now')),
                updated_at TEXT DEFAULT (datetime('now'))
            );";

        // ====================================================================
        // EJECUCIÓN PRINCIPAL
        // ====================================================================

        /// <summary>
        /// Ejecuta todos los pasos de la migración V24.
        ///
        /// ORDEN DE EJECUCIÓN:
        ///   1. Habilitar CR-SQLite en la conexión (requisito para crsql_as_crr).
        ///   2. Agregar columna subunidad a tablas existentes (idempotente).
        ///   3. Crear tablas nuevas: config, usuarios.
        ///   4. Crear tablas de control: sync_control, sync_archive.
        ///   5. Poblar subunidad en datos históricos.
        ///   6. Convertir tablas a CRDT con crsql_as_crr() (requiere commit previo).
        ///
        /// NOTA SOBRE TRANSACCIÓN:
        ///   crsql_as_crr() requiere que la transacción actual esté commiteada antes
        ///   de ejecutarse (la extensión necesita ver el esquema actualizado).
        ///   Por eso, pasos 2-5 van dentro de la transacción del SchemaMigrator,
        ///   y paso 6 (crsql_as_crr) se ejecuta DESPUÉS del commit del migrador.
        ///   WORKAROUND: Ejecutamos crsql_as_crr dentro de la transacción si CR-SQLite lo permite,
        ///   de lo contrario lo marcamos para post-commit.
        /// </summary>
        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // ---------------------------------------------------------------
            // Paso 1: Habilitar extensión CR-SQLite
            // ---------------------------------------------------------------
            var sqliteConn = connection as SQLiteConnection;
            bool crSqliteAvailable = false;
            if (sqliteConn != null)
            {
                crSqliteAvailable = CrSqliteLoader.EnableOnConnection(sqliteConn);
            }

            if (!crSqliteAvailable)
            {
                // La migración continúa sin CRDT. Los datos estarán disponibles
                // pero la sincronización no funcionará hasta que CR-SQLite esté disponible.
                LogMigration("ADVERTENCIA: CR-SQLite no disponible. Las tablas no se convertirán a CRDT. " +
                             "La sincronización estará deshabilitada.");
            }

            // ---------------------------------------------------------------
            // Paso 2: Agregar columna subunidad a tablas existentes
            // ---------------------------------------------------------------
            foreach (var (tableName, fallbackColumn) in ExistingTablesWithSubunidad)
            {
                AddSubunidadColumnIfMissing(connection, transaction, tableName);
            }

            // ---------------------------------------------------------------
            // Paso 3: Crear tablas nuevas (config, usuarios)
            // ---------------------------------------------------------------
            connection.Execute(CreateConfigSql, transaction: transaction);
            connection.Execute(CreateUsuariosSql, transaction: transaction);
            LogMigration("Tablas config y usuarios creadas/verificadas.");

            // Índice para búsqueda rápida por nombre de usuario Windows
            connection.Execute(
                "CREATE INDEX IF NOT EXISTS idx_usuarios_nombreusuario ON usuarios(nombreusuario);",
                transaction: transaction);

            // ---------------------------------------------------------------
            // Paso 4: Crear tablas de control de sync (sync_control, sync_archive)
            // ---------------------------------------------------------------
            var controlRepo = new SyncControlRepository();
            controlRepo.EnsureTablesExist(connection, transaction);
            LogMigration("Tablas de control sync_control y sync_archive creadas/verificadas.");

            // ---------------------------------------------------------------
            // Paso 5: Poblar subunidad en datos históricos
            // ---------------------------------------------------------------
            foreach (var (tableName, fallbackColumn) in ExistingTablesWithSubunidad)
            {
                PopulateSubunidadForExistingData(connection, transaction, tableName, fallbackColumn);
            }

            // ---------------------------------------------------------------
            // Paso 6: (Removido) Convertir tablas a CRDT
            // ---------------------------------------------------------------
            // La conversión a CRDT se realiza ahora en SyncInitializer.EnableAllCrdts()
            // porque no puede ejecutarse dentro de la transacción del SchemaMigrator.

            LogMigration("Migración V24 completada (esquema). CRDT se habilitarán en SyncInitializer.");
        }

        // ====================================================================
        // MÉTODOS AUXILIARES
        // ====================================================================

        /// <summary>
        /// Agrega la columna `subunidad TEXT DEFAULT ''` a una tabla si no existe.
        ///
        /// IDEMPOTENCIA: Lee pragma_table_info antes de ejecutar ALTER TABLE.
        /// SQLite no soporta ALTER TABLE ADD COLUMN IF NOT EXISTS.
        /// </summary>
        private void AddSubunidadColumnIfMissing(
            IDbConnection connection,
            IDbTransaction transaction,
            string tableName)
        {
            // Verificar si la tabla existe primero
            int tableExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=@TableName;",
                new { TableName = tableName },
                transaction: transaction);

            if (tableExists == 0)
            {
                LogMigration($"Tabla '{tableName}' no existe. Omitiendo ALTER TABLE.");
                return;
            }

            // Verificar si la columna subunidad ya existe
            string escapedName = tableName.Replace("'", "''");
            var columns = connection.Query(
                $"PRAGMA table_info('{escapedName}');",
                transaction: transaction).ToList();

            bool hasSubunidad = columns.Any(c =>
                string.Equals(c.name?.ToString(), "subunidad", StringComparison.OrdinalIgnoreCase));

            if (hasSubunidad)
            {
                LogMigration($"'{tableName}' ya tiene columna subunidad. Omitiendo ALTER TABLE.");
                return;
            }

            connection.Execute(
                $"ALTER TABLE [{tableName}] ADD COLUMN subunidad TEXT DEFAULT '';",
                transaction: transaction);

            LogMigration($"Columna subunidad agregada a '{tableName}'.");
        }

        /// <summary>
        /// Pobla la columna subunidad en registros históricos (pre-migración).
        ///
        /// ESTRATEGIA:
        ///   - Si la tabla tiene columna `departamento`: usar ese valor (terceros, sedes).
        ///   - Para el resto: dejar '' (vacío). El administrador debe actualizar manualmente
        ///     o se puede correr un script de corrección después.
        ///
        /// CONDICIÓN: Solo actualiza filas donde subunidad = '' (no sobreescribe datos existentes).
        /// </summary>
        private void PopulateSubunidadForExistingData(
            IDbConnection connection,
            IDbTransaction transaction,
            string tableName,
            string fallbackColumn)
        {
            if (string.IsNullOrEmpty(fallbackColumn)) { return; }

            // Verificar que la columna de fallback existe
            string escapedName = tableName.Replace("'", "''");
            var columns = connection.Query(
                $"PRAGMA table_info('{escapedName}');",
                transaction: transaction).ToList();

            bool hasFallback = columns.Any(c =>
                string.Equals(c.name?.ToString(), fallbackColumn, StringComparison.OrdinalIgnoreCase));

            if (!hasFallback) { return; }

            // Poblar subunidad usando la columna de departamento (convertida a mayúsculas)
            int updated = connection.Execute(
                $@"UPDATE [{tableName}]
                   SET subunidad = UPPER(COALESCE(TRIM([{fallbackColumn}]), ''))
                   WHERE (subunidad IS NULL OR subunidad = '')
                     AND [{fallbackColumn}] IS NOT NULL
                     AND [{fallbackColumn}] != '';",
                transaction: transaction);

            if (updated > 0)
            {
                LogMigration($"'{tableName}': {updated} registros históricos con subunidad='departamento'.");
            }
        }

        /// <summary>
        /// Helper de logging para mensajes de migración.
        /// Reutiliza LoggerHelper del proyecto.
        /// </summary>
        private static void LogMigration(string message)
        {
            Core.Helpers.LoggerHelper.Info($"[MigrationV24] {message}");
        }
    }
}
