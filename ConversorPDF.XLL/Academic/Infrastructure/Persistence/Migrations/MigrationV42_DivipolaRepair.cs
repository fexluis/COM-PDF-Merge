using System.Data;
using Dapper;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV42_DivipolaRepair : IMigration
    {
        public int Version => 42;
        public string Name => "Repair missing tables from bug v25";
        
        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // FASE 71 FIX: Crear si no existe usando el DDL canónico de 8 columnas (idéntico a Academic_divipola.sql).
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS divipola (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    departamento TEXT,
                    departamento_codigo TEXT,
                    departamento_normalizado TEXT,
                    municipio TEXT,
                    municipio_codigo TEXT,
                    municipio_normalizado TEXT,
                    municipio_categoria INTEGER DEFAULT 2
                );", transaction: transaction);
                
            // FASE 71 FIX: Verificar las columnas canónicas requeridas por AcademicRepository
            int faltantes = connection.QueryFirstOrDefault<int>(@"
                SELECT COUNT(*) FROM (
                    SELECT 'departamento_codigo' AS c UNION ALL SELECT 'departamento_normalizado'
                    UNION ALL SELECT 'municipio' UNION ALL SELECT 'municipio_codigo'
                    UNION ALL SELECT 'municipio_normalizado' UNION ALL SELECT 'municipio_categoria'
                ) req
                WHERE req.c NOT IN (SELECT name FROM pragma_table_info('divipola'));", transaction: transaction);

            if (faltantes > 0)
            {
                LoggerHelper.Warning($"[MigrationV42] divipola con esquema no canónico ({faltantes} columnas ausentes); la siembra (Fase 70) lo recreará.");
            }

            // Fase 26 FIX: Remover UNIQUE de CodigoTrayecto para no colisionar con lógica CRDT de SyncInitializer
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS BaseGastosTerrestre (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Territorial TEXT NOT NULL,
                    CodigoTrayecto TEXT NOT NULL,
                    Trayecto TEXT NOT NULL,
                    Valor REAL,
                    subunidad TEXT
                );", transaction: transaction);

            // Fase 59 y 61 FIX
            int terrColExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='BaseGastosTerrestre' AND sql LIKE '%Territorial%';", transaction: transaction);
            if (terrColExists == 0)
            {
                connection.Execute("ALTER TABLE BaseGastosTerrestre ADD COLUMN Territorial TEXT NOT NULL DEFAULT '';", transaction: transaction);
            }
            
            int subunidadBaseExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='BaseGastosTerrestre' AND sql LIKE '%subunidad%';", transaction: transaction);
            if (subunidadBaseExists == 0)
            {
                connection.Execute("ALTER TABLE BaseGastosTerrestre ADD COLUMN subunidad TEXT;", transaction: transaction);
            }

            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS territoriales (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    territorial TEXT, 
                    subunidad TEXT
                );", transaction: transaction);

            // Fase 59 y 61 FIX
            int subunidadExists = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='territoriales' AND sql LIKE '%subunidad%';", transaction: transaction);
            if (subunidadExists == 0)
            {
                connection.Execute("ALTER TABLE territoriales ADD COLUMN subunidad TEXT;", transaction: transaction);
            }

            // Fase 42 y 60 FIX: Repoblar data semilla canónica
            int terrCount = connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM territoriales;", transaction: transaction);
            if (terrCount == 0)
            {
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
        }
    }
}
