using System;
using System.Linq;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV30_AddTerminalAereo : IMigration
    {
        public int Version => 30;
        public string Name => "AddTerminalAereo";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // 1. Crear tabla terminales
            connection.Execute(@"
                CREATE TABLE IF NOT EXISTS terminales (
                    id INTEGER NOT NULL PRIMARY KEY,
                    terminal TEXT NOT NULL DEFAULT '',
                    maxvalor REAL NOT NULL DEFAULT 0
                );
            ", transaction: transaction);

            // 2. Insertar valores por defecto (Si la tabla está vacía, o ignorando duplicados si el id/nombre existe)
            // Se usa INSERT OR IGNORE usando el id (1 y 2) para asegurar que no duplique.
            connection.Execute(@"
                INSERT OR IGNORE INTO terminales (id, terminal, maxvalor)
                VALUES (1, 'El Dorado (Bogotá)', 500000);

                INSERT OR IGNORE INTO terminales (id, terminal, maxvalor)
                VALUES (2, 'José María Córdova (Medellín)', 450000);
            ", transaction: transaction);

            // 3. Agregar columnas a vinculaciones
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            var columns = connection.Query("PRAGMA table_info(vinculaciones)", transaction: transaction).Select(r => (string)((dynamic)r).name).ToList();
            if (!columns.Contains("isTerminalAereo", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN isTerminalAereo INTEGER NOT NULL DEFAULT 0;", transaction: transaction);
            if (!columns.Contains("idTerminales", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN idTerminales INTEGER NOT NULL DEFAULT 0;", transaction: transaction);
            if (!columns.Contains("gastoTerminalAerea", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN gastoTerminalAerea REAL NOT NULL DEFAULT 0;", transaction: transaction);
            
        }
    }
}
