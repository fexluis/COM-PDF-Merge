using System.Data;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// V83: Tabla de auditoría para la extensión o ajuste de ventanas temporales en 'periodos' (Fase 1854).
    /// Migración aditiva y tolerante: no muta datos existentes ni elimina tablas.
    /// </summary>
    public sealed class MigrationV83_AddAgPeriodoVentanaAuditoria : IMigration, IMigrationTolerante
    {
        public int Version => 83;
        public string Name => "AddAgPeriodoVentanaAuditoria";
        public bool Aditiva => true;

        /// <summary>
        /// Crea la tabla ag_periodo_ventana_auditoria y su índice para registrar cambios de ventana con actor y motivo.
        /// </summary>
        public void Up(IDbConnection connection, IDbTransaction transaction = null)
        {
            connection.Execute(@"CREATE TABLE IF NOT EXISTS ag_periodo_ventana_auditoria (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                idperiodo       INTEGER NOT NULL,
                periodo         TEXT,
                inicio_antes    TEXT,
                fin_antes       TEXT,
                inicio_despues  TEXT,
                fin_despues     TEXT,
                actor           TEXT NOT NULL,
                origen          TEXT NOT NULL,
                motivo          TEXT,
                fecha           TEXT NOT NULL
            );
            CREATE INDEX IF NOT EXISTS idx_ag_periodo_ventana_aud ON ag_periodo_ventana_auditoria(idperiodo);", transaction: transaction);
        }
    }
}
