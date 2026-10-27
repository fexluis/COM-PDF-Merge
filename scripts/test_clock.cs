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
            string[] tables = new[] { "salvedades_conflictos", "permisos" };
            foreach(var t in tables) {
                using (var cmd = new SQLiteCommand(string.Format("SELECT DISTINCT col_name FROM {0}__crsql_clock;", t), conn))
                using (var reader = cmd.ExecuteReader())
                {
                    Console.WriteLine("Cols in __crsql_clock for " + t + ":");
                    while(reader.Read()) {
                        Console.WriteLine("- " + reader[0]);
                    }
                }
            }
        }
    }
}
