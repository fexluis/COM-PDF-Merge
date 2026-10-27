using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración 43: Resolución de conflictos de esquema para Gastos Terrestres y Territoriales.
    /// Resuelve: 
    /// 1. Eliminación del constraint UNIQUE en CodigoTrayecto (bug originado en V18) recreando la tabla.
    /// 2. Inyección robusta de la columna 'subunidad' mediante PRAGMA table_info (bug originado en V24/V42).
    /// 3. Recreación/Seeding de la tabla territoriales (por exclusión de compilación de V42 en releases previas).
    /// </summary>
    public class MigrationV43_FixBaseGastosTerrestreSubunidad : IMigration
    {
        public int Version => 43;
        public string Name => "Añadir subunidad y remover UNIQUE en BaseGastosTerrestre";

        private void LogMigration(string message)
        {
            Core.Helpers.LoggerHelper.Info($"[MigrationV43] {message}");
        }

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (transaction == null)
                throw new InvalidOperationException("MigrationV43 requiere transacción explícita para garantizar atomicidad.");

            LogMigration("Iniciando revisión y reparación de BaseGastosTerrestre...");

            // 1. Revisar estado actual de BaseGastosTerrestre usando PRAGMA table_info
            var baseColumns = connection.Query(
                "SELECT name FROM pragma_table_info('BaseGastosTerrestre');",
                transaction: transaction).Select(c => (string)c.name).ToList();

            bool baseExists = baseColumns.Count > 0;
            bool baseHasSubunidad = baseColumns.Any(c => string.Equals(c, "subunidad", StringComparison.OrdinalIgnoreCase));

            // Verificar si tiene el constraint UNIQUE en CodigoTrayecto (u otros index únicos)
            int hasUnique = 0;
            if (baseExists)
            {
                hasUnique = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM pragma_index_list('BaseGastosTerrestre') WHERE origin IN ('u','c');",
                    transaction: transaction);
            }

            if (!baseExists)
            {
                // Crear desde cero si no existe
                LogMigration("BaseGastosTerrestre no existe. Creando nueva tabla...");
                connection.Execute(@"
                    CREATE TABLE BaseGastosTerrestre (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Territorial TEXT NOT NULL,
                        CodigoTrayecto TEXT NOT NULL,
                        Trayecto TEXT NOT NULL,
                        Valor REAL,
                        subunidad TEXT
                    );", transaction: transaction);
            }
            else if (hasUnique > 0 || !baseHasSubunidad)
            {
                // Si existe pero requiere cambios estructurales (UNIQUE persistente o falta subunidad)
                LogMigration($"BaseGastosTerrestre necesita reconstrucción. UNIQUE={hasUnique > 0}, subunidad={baseHasSubunidad}");
                
                // Construir lista de columnas a copiar (intersección entre vieja y nueva tabla)
                var oldColumns = string.Join(", ", baseColumns.Select(c => $"\"{c}\""));
                
                // FASE 63 FIX: Reconstrucción en dos pasos sin RENAME (CREATE tmp -> INSERT -> DROP original -> CREATE final -> INSERT -> DROP tmp)
                // Esto evita completamente 'ALTER TABLE ... RENAME TO' que en SQLite (>= 3.25) gatilla 
                // la re-evaluación/validación de todas las vistas del catálogo.
                connection.Execute(@"
                    CREATE TABLE BaseGastosTerrestre_tmp (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Territorial TEXT NOT NULL,
                        CodigoTrayecto TEXT NOT NULL,
                        Trayecto TEXT NOT NULL,
                        Valor REAL,
                        subunidad TEXT
                    );", transaction: transaction);

                // Validar que las columnas a insertar no contengan esquemas incompatibles (futuras migraciones aplicadas parcialmente)
                var newColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
                    { "Id", "Territorial", "CodigoTrayecto", "Trayecto", "Valor", "subunidad" };
                
                var unmapped = baseColumns.Where(c => !newColumns.Contains(c)).ToList();
                if (unmapped.Any())
                {
                    throw new InvalidOperationException(
                        $"MigrationV43: las columnas '{string.Join(", ", unmapped)}' de la tabla antigua no existen en la nueva. Esquema incompatible.");
                }

                // Insertar datos de la tabla antigua a la temporal y verificar cuenta
                int beforeCount = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM BaseGastosTerrestre;", transaction: transaction);
                LogMigration($"Registros antes de recrear: {beforeCount}");

                connection.Execute($@"
                    INSERT INTO BaseGastosTerrestre_tmp ({oldColumns}) 
                    SELECT {oldColumns} FROM BaseGastosTerrestre;", transaction: transaction);

                connection.Execute("DROP TABLE BaseGastosTerrestre;", transaction: transaction);

                connection.Execute(@"
                    CREATE TABLE BaseGastosTerrestre (
                        Id INTEGER PRIMARY KEY AUTOINCREMENT,
                        Territorial TEXT NOT NULL,
                        CodigoTrayecto TEXT NOT NULL,
                        Trayecto TEXT NOT NULL,
                        Valor REAL,
                        subunidad TEXT
                    );", transaction: transaction);

                connection.Execute($@"
                    INSERT INTO BaseGastosTerrestre ({oldColumns}) 
                    SELECT {oldColumns} FROM BaseGastosTerrestre_tmp;", transaction: transaction);

                connection.Execute("DROP TABLE BaseGastosTerrestre_tmp;", transaction: transaction);

                int afterCount = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM BaseGastosTerrestre;", transaction: transaction);
                if (afterCount != beforeCount)
                {
                    throw new InvalidOperationException(
                        $"V43: pérdida de datos al recrear BaseGastosTerrestre. Antes={beforeCount}, Después={afterCount}");
                }

                int postUnique = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM pragma_index_list('BaseGastosTerrestre') WHERE origin IN ('u','c');",
                    transaction: transaction);
                if (postUnique > 0)
                {
                    throw new InvalidOperationException("MigrationV43: constraint UNIQUE persiste tras recreación.");
                }

            }
            else
            {
                LogMigration("BaseGastosTerrestre está correcta. Sin cambios estructurales primarios.");
                
                var newCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
                    { "Id", "Territorial", "CodigoTrayecto", "Trayecto", "Valor", "subunidad" };
                
                var extras = baseColumns.Where(c => !newCols.Contains(c)).ToList();
                if (extras.Any())
                {
                    LogMigration($"ADVERTENCIA: Columnas no estándar detectadas ({string.Join(", ", extras)}). Verificar compatibilidad CRDT.");
                }
            }

            // Post-verificación estricta de BaseGastosTerrestre
            var verifyColumns = connection.Query(
                "SELECT name FROM pragma_table_info('BaseGastosTerrestre');",
                transaction: transaction).Select(c => (string)c.name).ToList();
            bool verified = verifyColumns.Any(c => string.Equals(c, "subunidad", StringComparison.OrdinalIgnoreCase));
            if (!verified)
            {
                throw new InvalidOperationException("MigrationV43: No se pudo agregar/verificar la columna 'subunidad' en BaseGastosTerrestre.");
            }

            // 2. Revisar y reparar tabla territoriales
            LogMigration("Verificando tabla territoriales...");
            var terrColumns = connection.Query(
                "SELECT name FROM pragma_table_info('territoriales');",
                transaction: transaction).Select(c => (string)c.name).ToList();

            if (terrColumns.Count == 0)
            {
                LogMigration("territoriales no existe. Creando tabla...");
                connection.Execute(@"
                    CREATE TABLE territoriales (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        territorial TEXT, 
                        subunidad TEXT
                    );", transaction: transaction);
            }
            else if (!terrColumns.Any(c => string.Equals(c, "subunidad", StringComparison.OrdinalIgnoreCase)))
            {
                LogMigration("Agregando subunidad a territoriales...");
                connection.Execute("ALTER TABLE territoriales ADD COLUMN subunidad TEXT;", transaction: transaction);
            }

            int terrCount = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM territoriales;", transaction: transaction);
            if (terrCount == 0)
            {
                LogMigration("territoriales vacía. Insertando data semilla canónica...");
                connection.Execute(@"
                    INSERT OR IGNORE INTO territoriales (id, territorial, subunidad) VALUES 
                    (1, 'ANTIOQUIA', 'ANTIOQUIA'),
                    (2, 'ATLÁNTICO - CESAR - MAGDALENA - LA GUAJIRA', 'ATLANTICO'),
                    (3, 'BOLIVAR - CÓRDOBA - SUCRE - SAN ANDRÉS', 'BOLIVAR'),
                    (4, 'BOYACÁ - CASANARE', 'BOYACA'),
                    (5, 'CALDAS', 'CALDAS'),
                    (6, 'CAUCA', 'CAUCA'),
                    (7, 'CHOCO', 'CHOCO'),
                    (8, 'CUNDINAMARCA', 'CUNDINAMARCA'),
                    (9, 'HUILA - CAQUETÁ - PUTUMAYO', 'HUILA'),
                    (10, 'META - GUAVIARE - GUAINÍA - VAUPÉS - VICHADA - AMAZONAS', 'META'),
                    (11, 'NARIÑO - ALTO PUTUMAYO', 'NARINO'),
                    (12, 'NORTE DE SANTANDER - ARAUCA', 'NORTE'),
                    (13, 'QUINDÍO - RISARALDA', 'RISARALDA'),
                    (14, 'SANTANDER', 'SANTANDER'),
                    (15, 'TOLIMA', 'TOLIMA'),
                    (16, 'VALLE', 'VALLE'),
                    (17, 'ESAP CENTRAL', 'ESAPCENTRAL');
                ", transaction: transaction);
            }

            LogMigration("Migración V43 finalizada exitosamente.");
        }
    }
}
