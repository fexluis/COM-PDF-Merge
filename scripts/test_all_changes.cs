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
                using(var cmd = new SQLiteCommand("SELECT DISTINCT \"table\" FROM crsql_changes", conn))
                using(var reader = cmd.ExecuteReader()) {
                    while(reader.Read()) {
                        Console.WriteLine(reader[0]);
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine("FAIL: " + ex.Message);
            }
        }
    }
}
