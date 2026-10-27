using System;
using System.Data.SQLite;
using System.IO;
using System.Text;

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
            StringBuilder sb = new StringBuilder();
            
            string[] tables = new[] { 
                "config", "usuarios", "salvedades_conflictos", "permisos", "divipola", "territoriales" 
            };
            foreach(var tbl in tables) {
                using (var cmd = new SQLiteCommand(string.Format("SELECT sql FROM sqlite_master WHERE name = '{0}' AND type='table';", tbl), conn))
                {
                    var sqlObj = cmd.ExecuteScalar();
                    var sql = sqlObj == null ? "" : sqlObj.ToString();
                    sb.AppendLine(string.Format("{0} -> {1}", tbl, sql));
                }
            }
            Console.WriteLine(sb.ToString());
        }
    }
}
