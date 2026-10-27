using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Fuerza la recreación de la tabla permisos en caso de anomalías en versiones previas.
    /// FASE 34: Desacoplamiento de sentencias (DROP, CREATE, INSERT) en llamadas separadas 
    /// a Execute() para prevenir fallos silenciosos en Dapper/SQLite al agrupar multi-statements.
    /// </summary>
    public class MigrationV37_RecreatePermisos : IMigration
    {
        public int Version => 37;
        public string Name => "RecreatePermisos";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // FASE 67 FIX: Verificar existencia y estructura de la tabla 'permisos' para no destruir
            // permisos personalizados en vaults que se re-migran durante una reparación de esquema.
            int exists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='permisos';", transaction: transaction);

            bool recreate = exists == 0;
            if (!recreate)
            {
                var cols = connection.Query("PRAGMA table_info(permisos);", transaction: transaction)
                                     .Select(c => (string)c.name).ToList();
                recreate = !cols.Contains("perfil", System.StringComparer.OrdinalIgnoreCase) 
                        || !cols.Contains("opcionespermitidas", System.StringComparer.OrdinalIgnoreCase);
            }

            if (recreate)
            {
                connection.Execute("DROP TABLE IF EXISTS permisos;", transaction: transaction);
                connection.Execute(@"
                    CREATE TABLE permisos (
                        perfil TEXT PRIMARY KEY NOT NULL,
                        opcionespermitidas TEXT NOT NULL
                    );
                ", transaction: transaction);
            }

            connection.Execute(@"
                INSERT OR IGNORE INTO permisos (perfil, opcionespermitidas) VALUES 
                ('Admin', 'ConvertirPDF,ConvertirObligacion,UnificarOB,UnificarOP,UnificarViaticos,MasivoViaticos,GenerarCumplidos,GenerarFilaSeleccionada,GestionDocumental,BuscarInformes,ProcesarFirma,SesionesAcademicas,Metadatos,Vinculaciones,Backups,Usuario,GestionUsuarios,Permisos'),
                ('Academica', 'ConvertirPDF,Vinculaciones,Backups,Usuario'),
                ('Contador', 'ConvertirPDF,ConvertirObligacion,Usuario'),
                ('Pagador', 'ConvertirPDF,UnificarOP,Usuario'),
                ('Auditor', 'ConvertirPDF,BuscarInformes,Usuario');
            ", transaction: transaction);
        }
    }
}
