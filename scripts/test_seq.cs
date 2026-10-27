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

            Console.WriteLine("sqlite_sequence:");
            using(var cmd = new SQLiteCommand("SELECT name, seq FROM sqlite_sequence", conn))
            using(var reader = cmd.ExecuteReader()) {
                while(reader.Read()) {
                    Console.WriteLine(string.Format("{0}: {1}", reader[0], reader[1]));
                }
            }
        }
    }
}
