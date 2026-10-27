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

            using(var cmdList = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%__crsql_clock';", conn))
            using(var readerList = cmdList.ExecuteReader())
            {
                while(readerList.Read())
                {
                    string clockTable = readerList["name"].ToString();
                    try {
                        using (var cmd = new SQLiteCommand(string.Format("SELECT DISTINCT col_name FROM {0};", clockTable), conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            Console.WriteLine(clockTable + ":");
                            while(reader.Read()) {
                                Console.WriteLine("  " + reader[0]);
                            }
                        }
                    } catch {}
                }
            }
        }
    }
}
