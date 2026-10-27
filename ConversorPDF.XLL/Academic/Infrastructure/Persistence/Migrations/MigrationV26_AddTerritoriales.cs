using System.Data;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V26: Crea y puebla la tabla territoriales en instalaciones existentes.
    /// Separa DDL (crear tabla) de DML (poblar filas) para idempotencia real:
    ///   - DDL: verificar sqlite_master antes de CREATE TABLE.
    ///   - DML: verificar COUNT(*) de filas antes de INSERT, garantizando que
    ///          tabla vacía por fallo parcial previo sea reparada.
    /// </summary>
    public class MigrationV26_AddTerritoriales : IMigration
    {
        public int Version => 26;
        public string Name => "Añade tabla territoriales para filtro de subunidades";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // 1. DDL idempotente — crear tabla solo si no existe
            int tableCount = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='territoriales';",
                transaction: transaction);

            if (tableCount == 0)
            {
                connection.Execute(@"
CREATE TABLE IF NOT EXISTS territoriales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    territorial TEXT,
    subunidad TEXT
);", transaction: transaction);
                LoggerHelper.Info("[MigrationV26] Tabla 'territoriales' creada (DDL).");
            }
            else
            {
                LoggerHelper.Info("[MigrationV26] Tabla 'territoriales' ya existe, omitiendo DDL.");
            }

            // 2. DML idempotente — insertar solo si tabla vacía
            int rowCount = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM territoriales;",
                transaction: transaction);

            if (rowCount == 0)
            {
                int inserted = connection.Execute(@"
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
(16, 'VALLE DEL CAUCA', 'VALLE'),
(17, 'ESAP CENTRAL', 'ESAPCENTRAL');", transaction: transaction);
                LoggerHelper.Info($"[MigrationV26] Tabla 'territoriales' poblada con {inserted} registros (DML).");
            }
            else
            {
                LoggerHelper.Info($"[MigrationV26] Tabla 'territoriales' ya tiene {rowCount} registros, omitiendo DML.");
            }
        }
    }
}
