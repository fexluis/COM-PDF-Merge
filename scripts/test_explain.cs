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
            
            using (var cmd = new SQLiteCommand("SELECT crsql_db_version();", conn)) {
                Console.WriteLine("crsql_db_version = " + cmd.ExecuteScalar());
            }
            
            // Try to find what crsql_changes is actually doing
            try {
                using (var cmd = new SQLiteCommand("EXPLAIN QUERY PLAN SELECT * FROM crsql_changes WHERE \"table\"='salvedades_conflictos'", conn)) 
                using (var reader = cmd.ExecuteReader()) {
                    while (reader.Read()) {
                        Console.WriteLine("EXPLAIN: " + reader["detail"]);
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine("EXPLAIN failed: " + ex.Message);
            }
        }
    }
}
