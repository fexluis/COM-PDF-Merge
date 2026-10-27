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
            
            try {
                using (var cmd = new SQLiteCommand("SELECT * FROM crsql_changes WHERE \"table\" = 'salvedades_conflictos' LIMIT 1;", conn))
                using (var reader = cmd.ExecuteReader())
                {
                    Console.WriteLine("OK: salvedades_conflictos");
                }
            } catch (Exception ex) {
                // Get the extended error message from the connection
                Console.WriteLine("FAIL: " + ex.Message);
            }
        }
    }
}
