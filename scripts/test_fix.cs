using System;
using System.Data.SQLite;
using System.IO;

class Program
{
    static void Main()
    {
        var tempDir = new DirectoryInfo(@"C:\Users\USUARIO\AppData\Local\Temp");
        var files = tempDir.GetFiles("esap_acad_tmp_AcademicData_BOLIVAR*.db");
        if (files.Length == 0) return;
        Array.Sort(files, (a, b) => b.LastWriteTime.CompareTo(a.LastWriteTime));
        var targetDb = files[0].FullName;
        
        using (var conn = new SQLiteConnection(string.Format("Data Source={0};Version=3;", targetDb)))
        {
            conn.Open();
            conn.EnableExtensions(true);
            conn.LoadExtension(@"C:\Users\USUARIO\AppData\Local\Temp\crsqlite_esap.dll", "sqlite3_crsqlite_init");
            
            // Fix salvedades_conflictos
            Console.WriteLine("Fixing salvedades_conflictos...");
            conn.Execute("DROP TABLE IF EXISTS salvedades_conflictos__crsql_clock;");
            conn.Execute("DROP TABLE IF EXISTS salvedades_conflictos__crsql_pks;");
            conn.Execute("DROP TRIGGER IF EXISTS salvedades_conflictos__crsql_dtrig;");
            conn.Execute("DROP TRIGGER IF EXISTS salvedades_conflictos__crsql_utrig;");
            conn.Execute("DROP TRIGGER IF EXISTS salvedades_conflictos__crsql_itrig;");
            conn.Execute("SELECT crsql_as_crr('salvedades_conflictos');");
            
            // Check if it works now
            try {
                using (var cmd2 = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'salvedades_conflictos' LIMIT 1;", conn))
                using (var reader2 = cmd2.ExecuteReader())
                {
                    Console.WriteLine("Query OK: salvedades_conflictos");
                }
            } catch (Exception ex) {
                Console.WriteLine("Error querying salvedades_conflictos: " + ex.Message);
            }
        }
    }
}
public static class SqliteExt {
    public static void Execute(this SQLiteConnection conn, string sql) {
        using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
    }
}
