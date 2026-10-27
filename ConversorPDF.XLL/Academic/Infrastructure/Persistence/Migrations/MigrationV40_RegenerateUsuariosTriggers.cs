using System.Data;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV40_RegenerateUsuariosTriggers : IMigration
    {
        public int Version => 40;
        public string Name => "RegenerateUsuariosTriggers";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // No hacemos cambios físicos en el esquema.
            // El único propósito de esta migración es incrementar la versión de la BD a 40.
            // Esto provocará que schemaMigrated sea TRUE en SyncInitializer, lo que activará
            // la lógica de auto-sanación (dropping and recreating triggers) para la tabla usuarios
            // y cualquier otra que haya quedado desfasada.
        }
    }
}
