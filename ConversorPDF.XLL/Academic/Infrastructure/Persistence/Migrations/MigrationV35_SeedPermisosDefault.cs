using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Crea la tabla de permisos e inserta los permisos por defecto.
    /// FASE 34: Desacoplamiento de sentencias (CREATE e INSERT) en llamadas separadas 
    /// a Execute() para prevenir fallos silenciosos en Dapper/SQLite al agrupar multi-statements.
    /// </summary>
    public class MigrationV35_SeedPermisosDefault : IMigration
    {
        public int Version => 35;
        public string Name => "SeedPermisosDefault";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Fase 13: Crear la tabla si no existe e insertar los permisos por defecto para Admin
            // Fase 13: Crear la tabla si no existe e insertar los permisos por defecto para Admin
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS permisos (
                    perfil TEXT PRIMARY KEY NOT NULL,
                    opcionespermitidas TEXT NOT NULL
                );
            ", transaction: transaction);

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
