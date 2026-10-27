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
            
            string[] tables = new[] {
                "salvedades_conflictos", "territoriales", "divipola", "permisos"
            };
            
            foreach(var table in tables) {
                try {
                    using (var cmd2 = new SQLiteCommand(string.Format("SELECT * FROM crsql_changes WHERE \"table\" = '{0}' LIMIT 1;", table), conn))
                    using (var reader2 = cmd2.ExecuteReader())
                    {
                        Console.WriteLine("Query OK: " + table);
                    }
                } catch (SQLiteException ex) {
                    Console.WriteLine(string.Format("Error querying {0}: ResultCode={1}, ExtendedErrCode={2}, Message={3}", table, ex.ResultCode, ex.ErrorCode, ex.Message));
                }
            }
        }
    }
}
