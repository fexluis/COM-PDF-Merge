using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Agrega el permiso de 'GestionUsuarios' al perfil Admin para bases de datos existentes
    /// que ya pasaron por la migración V37.
    /// </summary>
    public class MigrationV38_AddGestionUsuariosPermiso : IMigration
    {
        public int Version => 38;
        public string Name => "Agrega permiso de GestionUsuarios al perfil Admin";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // Append solo si falta; no pisa permisos custom del Admin
            connection.Execute(@"
                UPDATE permisos
                SET opcionespermitidas = TRIM(opcionespermitidas) || ',GestionUsuarios'
                WHERE perfil = 'Admin'
                  AND (',' || REPLACE(opcionespermitidas, ' ', '') || ',')
                      NOT LIKE '%,GestionUsuarios,%';
            ", transaction: transaction);
        }
    }
}
