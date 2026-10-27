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
            
            // Create a test table EXACTLY like salvedades_conflictos
            conn.Execute("CREATE TABLE test_salvedades ( id INTEGER PRIMARY KEY NOT NULL, idsesion_a INTEGER , idsesion_b INTEGER , idtercero INTEGER , idperiodo INTEGER, justificacion TEXT, fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP, subunidad TEXT DEFAULT '' );");
            conn.Execute("SELECT crsql_as_crr('test_salvedades');");
            
            try {
                using (var cmd2 = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'test_salvedades' LIMIT 1;", conn))
                using (var reader2 = cmd2.ExecuteReader())
                {
                    Console.WriteLine("Query OK: test_salvedades");
                }
            } catch (Exception ex) {
                Console.WriteLine("Error querying test_salvedades: " + ex.Message);
            }
            
            // Create a test table EXACTLY like permisos
            conn.Execute("CREATE TABLE test_permisos ( perfil TEXT PRIMARY KEY NOT NULL , opcionespermitidas TEXT  );");
            conn.Execute("SELECT crsql_as_crr('test_permisos');");
            
            try {
                using (var cmd2 = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'test_permisos' LIMIT 1;", conn))
                using (var reader2 = cmd2.ExecuteReader())
                {
                    Console.WriteLine("Query OK: test_permisos");
                }
            } catch (Exception ex) {
                Console.WriteLine("Error querying test_permisos: " + ex.Message);
            }
        }
    }
}
public static class SqliteExt {
    public static void Execute(this SQLiteConnection conn, string sql) {
        using (var cmd = new SQLiteCommand(sql, conn)) cmd.ExecuteNonQuery();
    }
}
