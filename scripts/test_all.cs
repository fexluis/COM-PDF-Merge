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
            
            using (var cmdList = new SQLiteCommand("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE '%__crsql%';", conn))
            using (var readerList = cmdList.ExecuteReader())
            {
                while (readerList.Read())
                {
                    string table = readerList["name"].ToString();
                    try {
                        using (var cmd = new SQLiteCommand(string.Format("SELECT * FROM crsql_changes WHERE \"table\" = '{0}' LIMIT 1;", table), conn))
                        using (var reader = cmd.ExecuteReader())
                        {
                            Console.WriteLine("OK: " + table);
                        }
                    } catch (Exception ex) {
                        Console.WriteLine("FAIL: " + table + " -> " + ex.Message);
                    }
                }
            }
        }
    }
}
