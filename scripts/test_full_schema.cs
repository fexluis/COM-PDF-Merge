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

            string[] tables = { "config", "salvedades_conflictos" };
            foreach(var t in tables) {
                Console.WriteLine("\n--- " + t + " ---");
                using(var cmd = new SQLiteCommand(string.Format("SELECT name, sql FROM sqlite_master WHERE name LIKE '{0}%'", t), conn))
                using(var reader = cmd.ExecuteReader()) {
                    while(reader.Read()) {
                        Console.WriteLine(reader[0] + ":\n" + reader[1]);
                    }
                }
            }
        }
    }
}
