using System.Data;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Helper para soltar triggers CR-SQLite obsoletos sobre la tabla 'periodos' antes de operaciones DDL/DML de migración.
    /// FASE 744 / FASE 819: Delega a CrdtTriggerRepair para reutilizar la lógica genérica de descarte de triggers CRDT.
    /// </summary>
    internal static class PeriodoVentanaCrdtRepair
    {
        /// <summary>
        /// Elimina los tres triggers CR-SQLite de la tabla periodos si existen.
        /// No altera la tabla __crsql_clock, preservando intacto el historial de sincronización multi-sede.
        /// Los triggers son regenerados limpiamente por SyncInitializer con el esquema actualizado.
        /// </summary>
        /// <param name="connection">Conexión activa a la base de datos SQLite.</param>
        /// <param name="transaction">Transacción activa en curso.</param>
        internal static void DropStaleCrsqlTriggers(IDbConnection connection, IDbTransaction transaction)
        {
            CrdtTriggerRepair.DropStaleTriggers(connection, transaction, "periodos");
        }
    }
}
