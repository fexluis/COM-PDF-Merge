using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Helper genérico para soltar triggers CR-SQLite obsoletos sobre cualquier tabla antes de operaciones DDL/DML de migración.
    /// FASE 819: Previene errores 'SQL logic error expected X values, got Y' causados por triggers con aridad de esquema previo.
    /// </summary>
    internal static class CrdtTriggerRepair
    {
        /// <summary>
        /// Elimina los tres triggers CR-SQLite de la tabla especificada si existen.
        /// No altera la tabla __crsql_clock, preservando intacto el historial de sincronización multi-sede.
        /// Los triggers son regenerados limpiamente por SyncInitializer con el esquema actualizado al completar las migraciones.
        /// </summary>
        /// <param name="connection">Conexión activa a la base de datos SQLite.</param>
        /// <param name="transaction">Transacción activa en curso.</param>
        /// <param name="tableName">Nombre de la tabla cuyos triggers CRDT deben soltarse.</param>
        internal static void DropStaleTriggers(IDbConnection connection, IDbTransaction transaction, string tableName)
        {
            if (string.IsNullOrWhiteSpace(tableName))
            {
                return;
            }

            string safeTable = tableName.Replace("'", "''").Replace("[", "").Replace("]", "");
            connection.Execute($"DROP TRIGGER IF EXISTS [{safeTable}__crsql_dtrig];", transaction: transaction);
            connection.Execute($"DROP TRIGGER IF EXISTS [{safeTable}__crsql_utrig];", transaction: transaction);
            connection.Execute($"DROP TRIGGER IF EXISTS [{safeTable}__crsql_itrig];", transaction: transaction);
        }
    }
}
