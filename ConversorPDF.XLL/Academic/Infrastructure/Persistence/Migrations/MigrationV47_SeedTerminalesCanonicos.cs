using System;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V47: Siembra y normaliza el catálogo canónico de 13 terminales aéreas en bases de datos existentes.
    /// Garantiza que el ID 1 sea 'NO APLICA' con maxvalor 0 y los IDs 2..13 contengan los aeropuertos y topes oficiales.
    /// Resuelve dinámicamente el ID de la terminal 'OTRO' para remapeo de vinculaciones legacy sin colisiones de IDs de usuario,
    /// y ajusta defensivamente los gastos que superen el nuevo maxvalor.
    /// </summary>
    public class MigrationV47_SeedTerminalesCanonicos : IMigration
    {
        public int Version => 47;
        public string Name => "SeedCanonicoTerminales";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='terminales';",
                    transaction: transaction) == 0)
            {
                return;
            }

            // Normalizar IDs legacy conocidos de V30
            connection.Execute(@"
                UPDATE terminales SET terminal = 'NO APLICA', maxvalor = 0 WHERE id = 1;
                UPDATE terminales SET terminal = 'ANTIOQUIA (Rionegro)', maxvalor = 162634 
                 WHERE id = 2 AND (terminal = 'José María Córdova (Medellín)' OR terminal = 'ANTIOQUIA (Rionegro)');
            ", transaction: transaction);

            // Fase 23 & Fase 31: Sembrar terminales canónicas faltantes evitando duplicar nombres existentes
            connection.Execute(@"
                INSERT OR IGNORE INTO terminales (id, terminal, maxvalor) VALUES
                (1, 'NO APLICA', 0),
                (2, 'ANTIOQUIA (Rionegro)', 162634),
                (3, 'ATLANTICO (Soledad)', 130704),
                (4, 'CORDOBA (Los Garzones)', 118731),
                (5, 'MAGDALENA (Santa Marta)', 129708),
                (6, 'NORTE DE SANTANDER', 50689),
                (7, 'QUINDIO (La Tebaida)', 186581),
                (8, 'SANTANDER (Lebrija)', 186581),
                (9, 'SUCRE (Corozal)', 162634),
                (10, 'VALLE DEL CAUCA (Palmira)', 186581),
                (11, 'NARIÑO (Chachagüí)', 186581),
                (12, 'PUTUMAYO (Pto.Asís)', 93788),
                (13, 'OTRO', 50689);
            ", transaction: transaction);

            // Fase 21 & Fase 32: Validar existencia de la tabla 'vinculaciones' y sus columnas antes de ejecutar actualizaciones
            bool vincExiste = connection.QueryFirstOrDefault<int>(
                "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';",
                transaction: transaction) > 0;

            if (vincExiste)
            {
                int colAereo = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM pragma_table_info('vinculaciones') WHERE name = 'isTerminalAereo';",
                    transaction: transaction);
                int colTerm = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM pragma_table_info('vinculaciones') WHERE name = 'idTerminales';",
                    transaction: transaction);
                int colGasto = connection.QueryFirstOrDefault<int>(
                    "SELECT COUNT(*) FROM pragma_table_info('vinculaciones') WHERE name = 'gastoTerminalAerea';",
                    transaction: transaction);

                if (colAereo > 0 && colTerm > 0)
                {
                    // Fase 26 & Fase 31: Resolver dinámicamente el ID de 'OTRO', prefiriendo el id canónico 13
                    int idOtro = connection.QueryFirstOrDefault<int>(
                        "SELECT id FROM terminales WHERE terminal = 'OTRO' ORDER BY (id = 13) DESC, id ASC LIMIT 1;",
                        transaction: transaction);

                    if (idOtro <= 0)
                    {
                        connection.Execute("INSERT INTO terminales (terminal, maxvalor) VALUES ('OTRO', 50689);", transaction: transaction);
                        idOtro = connection.QueryFirstOrDefault<int>("SELECT id FROM terminales WHERE terminal = 'OTRO' ORDER BY id DESC LIMIT 1;", transaction: transaction);
                    }

                    if (idOtro > 0)
                    {
                        connection.Execute(@"
                            UPDATE vinculaciones SET idTerminales = @IdOtro WHERE idTerminales = 1 AND isTerminalAereo = 1;
                        ", new { IdOtro = idOtro }, transaction: transaction);
                    }

                    // Fase 30 & Fase 34: Sanear gastos almacenados que superen el tope máximo de su terminal asociada para evitar bloqueos en UI
                    if (colGasto > 0)
                    {
                        connection.Execute(@"
                            UPDATE vinculaciones
                               SET gastoTerminalAerea = (SELECT t.maxvalor FROM terminales t WHERE t.id = vinculaciones.idTerminales)
                             WHERE isTerminalAereo = 1
                               AND EXISTS (SELECT 1 FROM terminales t
                                            WHERE t.id = vinculaciones.idTerminales AND t.maxvalor > 0)
                               AND gastoTerminalAerea > (SELECT t.maxvalor FROM terminales t WHERE t.id = vinculaciones.idTerminales);
                        ", transaction: transaction);
                    }
                }
            }
        }
    }
}
