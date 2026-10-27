using System;
using System.Data.SQLite;
using System.IO;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=== REPAIR SYNC CLOCK (CR-SQLite) ===");
        Console.WriteLine("Este script repara daños estructurales en __crsql_clock (SQL logic error) reconstruyendo los metadatos de sincronización de todas las tablas.");
        Console.WriteLine("ADVERTENCIA: Se perderá el historial de cambios locales pendientes de sincronizar (PUSH) en este nodo. Los datos locales permanecerán intactos.");

        // Intentar encontrar la base de datos desencriptada en Temp
        var tempDir = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Temp"));
        var files = tempDir.GetFiles("esap_acad_tmp_AcademicData_*.db").OrderByDescending(f => f.LastWriteTime).ToArray();

        if (files.Length == 0)
        {
            Console.WriteLine("\n[ERROR] No se encontró ninguna base de datos descifrada en Temp.");
            Console.WriteLine("Por favor, abra ConversorPDF en Excel para que el complemento extraiga la base de datos al directorio temporal, y mientras está abierto, ejecute este script.");
            return;
        }

        string targetDb = files[0].FullName;
        Console.WriteLine($"\n[INFO] Base de datos detectada: {targetDb}");
        Console.WriteLine("Presione ENTER para comenzar la reparación o Ctrl+C para cancelar...");
        Console.ReadLine();

        try
        {
            using (var conn = new SQLiteConnection($"Data Source={targetDb};Version=3;"))
            {
                conn.Open();
                
                // Cargar extensión crsqlite
                try 
                {
                    conn.EnableExtensions(true);
                    conn.LoadExtension(Path.Combine(tempDir.FullName, "crsqlite_esap.dll"), "sqlite3_crsqlite_init");
                    Console.WriteLine("[INFO] Extensión crsqlite_esap.dll cargada exitosamente.");
                } 
                catch(Exception ex) 
                {
                    Console.WriteLine($"[ERROR] No se pudo cargar crsqlite_esap.dll: {ex.Message}");
                    Console.WriteLine("Asegúrese de que el complemento haya extraído la DLL en Temp.");
                    return;
                }

                using (var transaction = conn.BeginTransaction())
                {
                    Console.WriteLine("[INFO] Reseteando tabla sync_control a 0.0...");
                    using (var cmd = new SQLiteCommand(@"
                        INSERT INTO sync_control (table_name, last_pushed_clock) 
                        VALUES ('__all__', '0.0')
                        ON CONFLICT(table_name) DO UPDATE SET last_pushed_clock = '0.0';", conn, transaction))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    Console.WriteLine("[INFO] Identificando tablas CRDT...");
                    var tablesQuery = @"
                        SELECT name FROM sqlite_master WHERE type='table' 
                        AND name NOT LIKE '%__crsql_%' AND name != 'sync_control'
                        AND name NOT LIKE '\_%' ESCAPE '\'";
                    
                    var tables = new System.Collections.Generic.List<string>();
                    using (var cmd = new SQLiteCommand(tablesQuery, conn, transaction))
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tables.Add(reader.GetString(0));
                        }
                    }

                    Console.WriteLine($"[INFO] Se encontraron {tables.Count} tablas base para reconstruir.");

                    foreach (var table in tables)
                    {
                        Console.WriteLine($"[INFO] Reconstruyendo metadata para: {table}");
                        
                        // Eliminar artifacts antiguos explícitamente
                        string safeTable = table.Replace("\"", "\"\"");
                        string[] dropStatements = {
                            $"DROP TABLE IF EXISTS \"{safeTable}__crsql_clock\";",
                            $"DROP TRIGGER IF EXISTS \"{safeTable}__crsql_dtrig\";",
                            $"DROP TRIGGER IF EXISTS \"{safeTable}__crsql_utrig\";",
                            $"DROP TRIGGER IF EXISTS \"{safeTable}__crsql_itrig\";"
                        };

                        foreach (var stmt in dropStatements)
                        {
                            using (var cmd = new SQLiteCommand(stmt, conn, transaction))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }

                        // Re-crear CRDT frescos
                        using (var cmd = new SQLiteCommand($"SELECT crsql_as_crr('{table}');", conn, transaction))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                    Console.WriteLine("\n[EXITO] Reparación estructural completada en la DB temporal.");
                    Console.WriteLine("Cierre este script. Luego, en Excel, interactúe con el complemento para forzar un re-cifrado (ej. guarde un cambio o cierre Excel).");
                    Console.WriteLine("En el próximo arranque, el PUSH se ejecutará limpiamente desde la versión 0.0 y los conflictos se resolverán por LWW.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n[ERROR FATAL] La reparación falló: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}
