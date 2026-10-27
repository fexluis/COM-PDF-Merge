using System.Data;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración sin carga DDL. Su único propósito es forzar el avance de versión 
    /// para que SqliteHelper devuelva schemaMigrated = true y desencadene la regeneración 
    /// masiva de triggers CR-SQLite en SyncInitializer.cs.
    /// (Reparación del error "SQL logic error expected 45 values, got 43")
    /// </summary>
    public class MigrationV32_RegenerateCrdtTriggers : IMigration
    {
        public int Version => 32;
        public string Name => "RegenerateCrdtTriggers";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Sin operación intencionalmente. 
            // La regeneración ocurre en SyncInitializer una vez esta migración se registra.
        }
    }
}
